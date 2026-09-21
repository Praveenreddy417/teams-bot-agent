import * as https from "https";
import { ClientSecretCredential } from "@azure/identity";
import { ConvRef } from "./Conversationstore";

const GRAPH_APP_ID = process.env.TEAMS_APP_ID ?? ""; // manifest/externalId — NOT the catalog internal id, NOT bot CLIENT_ID
const BOT_APP_ID = process.env.CLIENT_ID ?? "";
const DEFAULT_SERVICE_URL = process.env.BOT_SERVICE_URL ?? "https://smba.trafficmanager.net/in/";

// Optional config so we never need AppCatalog.Read.All:
//   TEAMS_CATALOG_APP_ID  — the app's internal catalog id (fastest; zero extra Graph calls)
//   TEAMS_SEED_TEAM_ID    — a team where the app is ALREADY installed; we read the catalog id from it
//   TEAMS_APP_NAME        — fallback display-name match if externalId isn't returned by the expand
const CONFIG_CATALOG_APP_ID = process.env.TEAMS_CATALOG_APP_ID ?? "";
const SEED_TEAM_ID = process.env.TEAMS_SEED_TEAM_ID ?? "";
const APP_DISPLAY_NAME = process.env.TEAMS_APP_NAME ?? "ESI Boomi Integration Agent";

function httpsJson<T = any>(
    method: string,
    url: string,
    token: string,
    body?: unknown
): Promise<{ status: number; data: T }> {
    return new Promise((resolve, reject) => {
        const u = new URL(url);
        const bodyBuf = body !== undefined ? Buffer.from(JSON.stringify(body), "utf-8") : undefined;
        const req = https.request(
            {
                hostname: u.hostname,
                path: u.pathname + u.search,
                method,
                headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${token}`,
                    ...(bodyBuf ? { "Content-Length": bodyBuf.byteLength } : {}),
                },
            },
            (res) => {
                let raw = "";
                res.on("data", (c) => (raw += c));
                res.on("end", () => {
                    let parsed: any = {};
                    try { parsed = raw ? JSON.parse(raw) : {}; } catch { parsed = { raw }; }
                    resolve({ status: res.statusCode ?? 0, data: parsed });
                });
            }
        );
        req.on("error", reject);
        if (bodyBuf) req.write(bodyBuf);
        req.end();
    });
}

async function getGraphToken(): Promise<string> {
    const cred = new ClientSecretCredential(
        process.env.TENANT_ID!,
        process.env.CLIENT_ID!,
        process.env.CLIENT_SECRET!
    );
    const resp = await cred.getToken("https://graph.microsoft.com/.default");
    if (!resp?.token) throw new Error("Empty Graph token");
    return resp.token;
}

async function getBotFrameworkToken(): Promise<string> {
    const cred = new ClientSecretCredential(
        process.env.TENANT_ID!,
        process.env.CLIENT_ID!,
        process.env.CLIENT_SECRET!
    );
    const resp = await cred.getToken("https://api.botframework.com/.default");
    if (!resp?.token) throw new Error("Empty Bot Framework token");
    return resp.token;
}

let cachedCatalogAppId: string | null = CONFIG_CATALOG_APP_ID || null;
// Catalog ids that failed to bind (app definition doesn't exist for them) — never reuse or
// re-discover these; a sideloaded-only app can keep surfacing the same bad id on re-scan.
const badCatalogAppIds = new Set<string>();

// A failed tenant scan is expensive (one Graph call per team). Remember the failure for a
// while so repeated alerts don't re-scan every team on every call.
let lastFailedScanAt = 0;
const FAILED_SCAN_COOLDOWN_MS = 5 * 60_000;
// De-dupe concurrent scans — parallel alerts would otherwise each start their own.
let inFlightScan: Promise<string | null> | null = null;

/** Match our app inside a team's installedApps list (by externalId, else display name). */
function matchInstalledApp(value: any[]): any | undefined {
    if (!Array.isArray(value)) return undefined;
    return value.find(
        (a) =>
            a?.teamsApp?.externalId === GRAPH_APP_ID ||
            (APP_DISPLAY_NAME && a?.teamsApp?.displayName === APP_DISPLAY_NAME)
    );
}

/**
 * Read a team's installed apps (needs TeamsAppInstallation.ReadWriteForTeam.All — NOT AppCatalog.Read.All)
 * and return our app's catalog id if it's installed there, else null.
 */
async function readCatalogAppIdFromTeam(token: string, teamId: string): Promise<string | null> {
    const list = await httpsJson(
        "GET",
        `https://graph.microsoft.com/v1.0/teams/${teamId}/installedApps?$expand=teamsApp`,
        token
    );
    if (list.data?.error) {
        console.warn(`[Channel] installedApps read failed for team ${teamId}: ${JSON.stringify(list.data.error)}`);
        return null;
    }
    const mine = matchInstalledApp(list.data?.value);
    return mine?.teamsApp?.id ?? null;
}

/**
 * Last resort: scan the tenant's teams for one that already has our app installed and read the
 * catalog id off it. Uses Group.ReadWrite.All (list teams) + TeamsAppInstallation.ReadWriteForTeam.All
 * (read installedApps) — no AppCatalog.Read.All. Result is cached for the process lifetime.
 */
async function discoverCatalogAppIdByScanningTeams(token: string): Promise<string | null> {
    const teamsResp = await httpsJson(
        "GET",
        "https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName&$top=100",
        token
    );
    if (teamsResp.data?.error) {
        console.warn(`[Channel] Team scan failed: ${JSON.stringify(teamsResp.data.error)}`);
        return null;
    }
    const teams: any[] = Array.isArray(teamsResp.data?.value) ? teamsResp.data.value : [];
    console.log(`[Channel] Scanning ${teams.length} team(s) for an existing install of "${APP_DISPLAY_NAME}"...`);

    // Small parallel batches so a large tenant doesn't serialise into a long wait.
    const BATCH = 8;
    for (let i = 0; i < teams.length; i += BATCH) {
        const batch = teams.slice(i, i + BATCH);
        const found = await Promise.all(
            batch.map((t) => readCatalogAppIdFromTeam(token, t.id).catch(() => null))
        );
        const hitIndex = found.findIndex((id) => !!id && !badCatalogAppIds.has(id));
        if (hitIndex >= 0) {
            const hit = found[hitIndex]!;
            console.log(`[Channel] Discovered catalog app id ${hit} from team "${batch[hitIndex].displayName}".`);
            return hit;
        }
    }
    return null;
}

/** Resolve the catalog app id without AppCatalog.Read.All: env → seed team → tenant scan. */
async function resolveCatalogAppId(token: string): Promise<string> {
    if (cachedCatalogAppId && !badCatalogAppIds.has(cachedCatalogAppId)) return cachedCatalogAppId;

    if (SEED_TEAM_ID) {
        const fromSeed = await readCatalogAppIdFromTeam(token, SEED_TEAM_ID);
        if (fromSeed && !badCatalogAppIds.has(fromSeed)) {
            cachedCatalogAppId = fromSeed;
            console.log(`[Channel] Resolved catalog app id from seed team: ${fromSeed}`);
            return fromSeed;
        }
    }

    const sinceLastFail = Date.now() - lastFailedScanAt;
    if (sinceLastFail < FAILED_SCAN_COOLDOWN_MS) {
        throw new Error(
            `Catalog id for "${APP_DISPLAY_NAME}" is still unknown — a tenant-wide scan failed ` +
            `${Math.round(sinceLastFail / 1000)}s ago and is not repeated for ` +
            `${FAILED_SCAN_COOLDOWN_MS / 60_000} minutes. Install the app into one team via Teams, ` +
            `or set TEAMS_CATALOG_APP_ID / TEAMS_SEED_TEAM_ID, then restart.`
        );
    }

    // Share a single scan across concurrent callers.
    if (!inFlightScan) {
        inFlightScan = discoverCatalogAppIdByScanningTeams(token).finally(() => {
            inFlightScan = null;
        });
    }
    const discovered = await inFlightScan;
    if (discovered) {
        cachedCatalogAppId = discovered;
        lastFailedScanAt = 0;
        return discovered;
    }
    lastFailedScanAt = Date.now();

    throw new Error(
        `Could not find "${APP_DISPLAY_NAME}" installed in any team, so its catalog id is unknown ` +
        `(reading the app catalog directly needs AppCatalog.Read.All, which is not granted). ` +
        `Fix (any one): install the app into at least ONE team via Teams (Manage team → Apps → Add an app) ` +
        `and it will be auto-discovered from then on; or set TEAMS_CATALOG_APP_ID; or set TEAMS_SEED_TEAM_ID. ` +
        `If the app IS installed somewhere, set TEAMS_APP_NAME to its exact display name in Teams.`
    );
}

/**
 * Ensure the bot's Teams app is installed on the given team.
 * Returns true if it performed a FRESH install (caller should wait for roster propagation).
 * Uses only TeamsAppInstallation.ReadWriteForTeam.All.
 */
export async function ensureAppInstalledInTeam(teamId: string): Promise<boolean> {
    if (!GRAPH_APP_ID && !CONFIG_CATALOG_APP_ID) {
        throw new Error("TEAMS_APP_ID (or TEAMS_CATALOG_APP_ID) env var not set — needed to install app into team");
    }
    const token = await getGraphToken();

    // 1. Already installed in this team? Reading it also yields the catalog id — no AppCatalog needed.
    const existing = await readCatalogAppIdFromTeam(token, teamId);
    if (existing) {
        cachedCatalogAppId = existing; // cache for future installs into other teams
        return false;
    }

    // 2. Not installed here — need the catalog id to bind the install. One retry: a catalog id
    // discovered from a team where the app was only ever sideloaded (never published to the org
    // app catalog) binds to nothing and 404s — evict it and try to resolve a different one.
    for (let attempt = 0; attempt < 2; attempt++) {
        const catalogAppId = await resolveCatalogAppId(token);

        const install = await httpsJson(
            "POST",
            `https://graph.microsoft.com/v1.0/teams/${teamId}/installedApps`,
            token,
            { "teamsApp@odata.bind": `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${catalogAppId}` }
        );
        if (install.status === 409) return false; // already installed (race)
        if (install.status === 201 || install.status === 200) return true;

        const isMissingAppDefinition =
            install.status === 404 &&
            JSON.stringify(install.data ?? "").includes("non existent App definition");

        if (isMissingAppDefinition) {
            badCatalogAppIds.add(catalogAppId);
            if (cachedCatalogAppId === catalogAppId) cachedCatalogAppId = null;
            if (attempt === 0) continue; // one retry with a freshly resolved id
            throw new Error(
                `Could not install "${APP_DISPLAY_NAME}" into team ${teamId}: every discoverable catalog id ` +
                `(incl. ${catalogAppId}) 404s with "non existent App definition". This means the app has never ` +
                `been published to the tenant's org-wide Teams app catalog — it's only been sideloaded per-team, ` +
                `and Graph's /appCatalogs/teamsApps/{id} bind only works for org-catalog apps. ` +
                `Fix: have a Teams admin upload the app package to Teams Admin Center (Manage apps → Upload new app), ` +
                `then set TEAMS_CATALOG_APP_ID to that catalog entry's id.`
            );
        }

        throw new Error(`Failed to install app in team ${teamId}: ${install.status} ${JSON.stringify(install.data)}`);
    }
    return true;
}

/**
 * Post a card into a Teams channel.
 *
 * Strategy 1 (primary): POST straight to /v3/conversations/{channelId}/activities.
 * In Teams a channel id IS a usable conversation id, so this starts a new thread in the
 * channel without needing POST /v3/conversations (which enforces a roster check) and
 * without needing any Microsoft Graph call or app-catalog permission at all.
 *
 * Strategy 2 (fallback): best-effort Graph install, then POST /v3/conversations with the
 * card as the initial activity — this both creates the conversation and posts the message.
 */
export async function postCardToChannel(
    teamId: string,
    channelId: string,
    cardJson: string,
    tenantId?: string
): Promise<{ conversationId: string; strategy: string }> {
    const botToken = await getBotFrameworkToken();
    const card = JSON.parse(cardJson);

    // ── Strategy 1: direct post to the channel ────────────────────────────────
    const direct = await httpsJson<{ id?: string }>(
        "POST",
        `${DEFAULT_SERVICE_URL}v3/conversations/${encodeURIComponent(channelId)}/activities`,
        botToken,
        card
    );
    if (direct.status >= 200 && direct.status < 300) {
        return { conversationId: channelId, strategy: "direct" };
    }
    console.warn(
        `[Channel] Direct post to ${channelId} failed (${direct.status}): ${JSON.stringify(direct.data)} — falling back to conversation create.`
    );

    // ── Strategy 2: ensure installed (best effort), then create + post ────────
    let freshlyInstalled = false;
    let installError: string | null = null;
    try {
        freshlyInstalled = await ensureAppInstalledInTeam(teamId);
        console.log(`[Channel] App install check for team ${teamId}: ${freshlyInstalled ? "installed" : "already present"}`);
        if (freshlyInstalled) {
            console.log(`[Channel] Freshly installed app in team ${teamId} — waiting for roster to propagate...`);
            await new Promise((r) => setTimeout(r, 5000));
        }
    } catch (err: any) {
        // Missing Graph permissions must not block the send — strategy 2 may still work
        // if the app was installed manually. Keep the reason so it surfaces in the final error
        // instead of silently collapsing into the same generic BotNotInConversationRoster.
        installError = err.message;
        console.warn(`[Channel] Skipping Graph install for team ${teamId}: ${err.message}`);
    }

    const payload = {
        channelData: {
            channel: { id: channelId },
            team: { id: teamId },
            tenant: tenantId ? { id: tenantId } : undefined,
        },
        isGroup: true,
        bot: { id: BOT_APP_ID },
        activity: card, // creating the conversation also delivers the card
    };

    let lastErr: any = direct;
    for (let attempt = 0; attempt < 4; attempt++) {
        const resp = await httpsJson<{ id?: string }>(
            "POST",
            `${DEFAULT_SERVICE_URL}v3/conversations`,
            botToken,
            payload
        );
        if (resp.status >= 200 && resp.status < 300 && resp.data?.id) {
            return { conversationId: resp.data.id, strategy: "conversation-create" };
        }
        lastErr = resp;
        const isRoster = JSON.stringify(resp.data ?? "").includes("BotNotInConversationRoster");
        await new Promise((r) => setTimeout(r, (isRoster ? 5000 : 1500) * (attempt + 1)));
    }

    const installOutcome = installError
        ? `failed — ${installError}`
        : freshlyInstalled
            ? "freshly installed, roster may still be propagating"
            : "already present";

    throw new Error(
        `Could not post to channel ${channelId}. Direct post: ${direct.status} ${JSON.stringify(direct.data)}. ` +
        `App install: ${installOutcome}. ` +
        `Conversation create: ${JSON.stringify(lastErr?.data)}`
    );
}

/**
 * Proactively resolve a Bot Framework conversation reference for a channel — LIVE, every call.
 * Installs the bot's app into the team first (if needed), then opens a channel conversation.
 */
export async function ensureChannelConversationRef(
    teamId: string,
    channelId: string,
    tenantId?: string
): Promise<ConvRef> {
    const freshlyInstalled = await ensureAppInstalledInTeam(teamId);

    // A just-installed app takes several seconds to appear in the channel roster; the
    // Bot Framework returns BotNotInConversationRoster until it does. Give it a head start.
    if (freshlyInstalled) {
        console.log(`[Channel] Freshly installed app in team ${teamId} — waiting for roster to propagate...`);
        await new Promise((r) => setTimeout(r, 5000));
    }

    const botToken = await getBotFrameworkToken();
    const payload = {
        channelData: {
            channel: { id: channelId },
            team: { id: teamId },
            tenant: tenantId ? { id: tenantId } : undefined,
        },
        isGroup: true,
        bot: { id: BOT_APP_ID },
        activity: {
            type: "message",
            text: "Bot connected for alerting.",
        },
    };

    // Retry with backoff: roster propagation after a fresh install can take 10–30s.
    let lastErr: any;
    for (let attempt = 0; attempt < 6; attempt++) {
        const resp = await httpsJson<{ id?: string; activityId?: string }>(
            "POST",
            `${DEFAULT_SERVICE_URL}v3/conversations`,
            botToken,
            payload
        );
        if (resp.status >= 200 && resp.status < 300 && resp.data?.id) {
            return {
                conversationId: resp.data.id,
                conversationType: "channel",
                channelId,
                teamId,
                serviceUrl: DEFAULT_SERVICE_URL,
                tenantId: tenantId ?? "",
                botId: BOT_APP_ID,
                userId: "",
                userName: "",
                userEmail: "",
            } as ConvRef;
        }
        lastErr = resp;
        const isRoster = JSON.stringify(resp.data ?? "").includes("BotNotInConversationRoster");
        await new Promise((r) => setTimeout(r, (isRoster ? 5000 : 1500) * (attempt + 1)));
    }
    throw new Error(`Could not create channel conversation for ${channelId}: ${JSON.stringify(lastErr?.data)}`);
}
