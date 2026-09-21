import * as http from "http";
import * as dotenv from "dotenv";
import { ensureChannelConversationRef, ensureAppInstalledInTeam } from "./Channelinstallservice";
import { greetAndRegisterRoster } from "./Proactivegreetservice";
dotenv.config({ path: "./env/.env.local" });

// ── Graph app-only token (client credentials) ──────────────────────
let cachedToken: { token: string; expiresAt: number } | null = null;

async function getGraphToken(): Promise<string> {
    if (cachedToken && cachedToken.expiresAt > Date.now() + 60_000) {
        return cachedToken.token;
    }

    const tenantId = process.env.AAD_TENANT_ID ?? process.env.TENANT_ID!;
    const clientId = process.env.AAD_CLIENT_ID ?? process.env.CLIENT_ID!;
    const clientSecret = process.env.AAD_CLIENT_SECRET ?? process.env.CLIENT_SECRET!;

    const params = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
    });

    const res = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: params.toString(),
        }
    );

    if (!res.ok) {
        throw new Error(`Graph token fetch failed: ${res.status} ${await res.text()}`);
    }

    const data = await res.json();
    cachedToken = { token: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
    return cachedToken.token;
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

// ── Types ────────────────────────────────────────────────────────────
export interface TeamMemberInput {
    userId: string;
    role: "owner" | "member";
}

export interface ChannelMemberInput {
    userId: string;
    role: "owner" | "member" | "guest";
}

export interface CreateTeamChannelRequest {
    teamDisplayName: string;
    teamDescription?: string;
    teamMembers: TeamMemberInput[];

    channelDisplayName: string;
    channelDescription?: string;
    membershipType?: "standard" | "private" | "shared";
    channelMembers?: ChannelMemberInput[];
}

export interface CreateChannelResult {
    status: 201 | 202;
    channel?: any;
    operationLocation?: string | null;
}

interface CreateTeamResult {
    teamId: string;
    failedMembers: string[];
}

// ── Step 1: create a new Team (async — must poll) ─────────────────────
async function createTeam(
    displayName: string,
    description: string,
    members: TeamMemberInput[]
): Promise<CreateTeamResult> {
    const token = await getGraphToken();

    const owners = members.filter((m) => m.role === "owner");
    if (owners.length < 1) {
        throw new Error("At least one owner is required in teamMembers");
    }
    // Graph's team-creation call accepts only one owner in the initial body;
    // any additional owners are added afterward via the members loop below.
    const owner = owners[0];

    // Pre-validate that the owner exists to prevent 404 template failures
    const userCheck = await fetch(`https://graph.microsoft.com/v1.0/users/${owner.userId}`, {
        headers: { Authorization: `Bearer ${token}` }
    });
    if (!userCheck.ok) {
        throw new Error(`Designated team owner (${owner.userId}) was not found in the directory.`);
    }

    const body = {
        "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
        displayName,
        description,
        members: [
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${owner.userId}')`,
                roles: ["owner"],
            },
        ],
    };

    const res = await fetch("https://graph.microsoft.com/v1.0/teams", {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(body),
    });

    if (res.status !== 202) {
        throw new Error(`Team creation failed: ${res.status} ${await res.text()}`);
    }

    const operationLocation = res.headers.get("location") ?? res.headers.get("content-location");
    if (!operationLocation) throw new Error("No operation location returned for team creation");

    let teamId: string | undefined;
    for (let attempt = 0; attempt < 30; attempt++) {
        await sleep(4000);

        const absoluteUrl = operationLocation.startsWith("http")
            ? operationLocation
            : `https://graph.microsoft.com/v1.0${operationLocation}`;

        const opRes = await fetch(absoluteUrl, { headers: { Authorization: `Bearer ${token}` } });
        const opData = await opRes.json();

        if (opData.status === "succeeded") {
            teamId = opData.targetResourceId;
            break;
        }
        if (opData.status === "failed") {
            throw new Error(`Team provisioning failed: ${JSON.stringify(opData)}`);
        }
    }
    if (!teamId) throw new Error("Team provisioning timed out after ~2 minutes");

    // ── Add remaining team members now that the team exists ──
    const failedMembers: string[] = [];
    const rest = members.filter((m) => m.userId !== owner.userId);
    for (const m of rest) {
        const addRes = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/members`, {
            method: "POST",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${m.userId}')`,
                roles: m.role === "owner" ? ["owner"] : [],
            }),
        });
        if (!addRes.ok) {
            console.error(`Failed to add team member ${m.userId}: ${addRes.status} ${await addRes.text()}`);
            failedMembers.push(m.userId);
        }
    }

    return { teamId, failedMembers };
}

// ── Step 2: create channel inside the (now-existing) team ─────────────
async function createDynamicChannel(
    teamId: string,
    displayName: string,
    description: string,
    membershipType: "standard" | "private" | "shared",
    members?: ChannelMemberInput[]
): Promise<CreateChannelResult> {
    const token = await getGraphToken();

    const body: Record<string, any> = { displayName, description, membershipType };

    if (membershipType !== "standard" && members?.length) {
        const list =
            membershipType === "shared"
                ? members.filter((m) => m.role === "owner").slice(0, 1)
                : members;

        body["@odata.type"] = "#Microsoft.Graph.channel";
        body.members = list.map((m) => ({
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${m.userId}')`,
            roles: [m.role],
        }));
    }

    const res = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(body),
    });

    if (res.status === 201) return { status: 201, channel: await res.json() };
    if (res.status === 202)
        return {
            status: 202,
            operationLocation: res.headers.get("location") ?? res.headers.get("content-location"),
        };

    throw new Error(`Channel creation failed: ${res.status} ${await res.text()}`);
}

// ── Retry wrapper — Graph sometimes 404s on a freshly-provisioned team
async function createDynamicChannelWithRetry(
    teamId: string,
    displayName: string,
    description: string,
    membershipType: "standard" | "private" | "shared",
    members: ChannelMemberInput[] | undefined,
    maxAttempts = 5,
    delayMs = 5000
): Promise<CreateChannelResult> {
    let lastErr: any;
    for (let i = 0; i < maxAttempts; i++) {
        try {
            return await createDynamicChannel(teamId, displayName, description, membershipType, members);
        } catch (err) {
            lastErr = err;
            if (i === maxAttempts - 1) break;
            await sleep(delayMs);
        }
    }
    throw lastErr;
}

// ── Resolve a channel created asynchronously (Graph returns 202 for shared channels) ──
// Without this, `result.channel` is undefined and the response omits channelId entirely.
async function resolveAsyncChannel(teamId: string, operationLocation: string): Promise<any | null> {
    const token = await getGraphToken();

    for (let attempt = 0; attempt < 15; attempt++) {
        await sleep(4000);

        const absoluteUrl = operationLocation.startsWith("http")
            ? operationLocation
            : `https://graph.microsoft.com/v1.0${operationLocation}`;

        const opRes = await fetch(absoluteUrl, { headers: { Authorization: `Bearer ${token}` } });
        if (!opRes.ok) continue;

        const opData = await opRes.json();
        if (opData.status === "failed") {
            throw new Error(`Channel provisioning failed: ${JSON.stringify(opData)}`);
        }
        if (opData.status !== "succeeded") continue;

        const channelId = opData.targetResourceId;
        if (!channelId) return null;

        const chRes = await fetch(
            `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}`,
            { headers: { Authorization: `Bearer ${token}` } }
        );
        return chRes.ok ? await chRes.json() : { id: channelId };
    }
    return null;
}

// ── Add a channel member after creation (used for "shared" channels only) ──
async function addChannelMember(teamId: string, channelId: string, member: ChannelMemberInput) {
    const token = await getGraphToken();
    const res = await fetch(
        `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/members`,
        {
            method: "POST",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${member.userId}')`,
                roles: [member.role],
            }),
        }
    );
    if (!res.ok) throw new Error(`Add member failed: ${res.status} ${await res.text()}`);
    return res.json();
}

// ── Helper: read raw Node request body (no express/body-parser) ────────
function readJsonBody(req: http.IncomingMessage): Promise<any> {
    return new Promise((resolve, reject) => {
        let raw = "";
        req.on("data", (c) => (raw += c));
        req.on("end", () => {
            try { resolve(raw ? JSON.parse(raw) : {}); }
            catch { reject(new Error("Invalid JSON body")); }
        });
        req.on("error", reject);
    });
}

// ── Plain Node handler: GET /api/teamsbot-list-teams ───────────────────
export async function handleListTeams(req: http.IncomingMessage, res: http.ServerResponse) {
    res.setHeader("Content-Type", "application/json");
    try {
        const token = await getGraphToken();
        const result = await fetch(
            "https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description",
            { headers: { Authorization: `Bearer ${token}` } }
        );
        if (!result.ok) throw new Error(`${result.status} ${await result.text()}`);
        const data = await result.json();
        res.writeHead(200);
        res.end(JSON.stringify({
            teams: data.value.map((t: any) => ({
                teamId: t.id,
                displayName: t.displayName,
                description: t.description,
            })),
        }));
    } catch (err: any) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
    }
}

// ── Plain Node handler: POST /api/teamsbot-greet-team ─────────────────────────
// Backfill for teams where the app was ALREADY installed (so the install event has
// long passed). Greets + registers every member so /api/alert/members works for them.
// Body: { teamId } or { channelId }  (+ optional skipExisting, default true)
export async function handleGreetTeam(req: http.IncomingMessage, res: http.ServerResponse) {
    res.setHeader("Content-Type", "application/json");

    let body: { teamId?: string; channelId?: string; skipExisting?: boolean };
    try {
        body = await readJsonBody(req);
    } catch {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "Invalid JSON body" }));
        return;
    }

    if (!body.teamId && !body.channelId) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "teamId or channelId is required" }));
        return;
    }

    try {
        // The Bot Framework roster is keyed by the channel thread id (19:...@thread.tacv2),
        // so translate a team GUID into its General channel first.
        let conversationId = body.channelId;
        if (!conversationId) {
            const token = await getGraphToken();
            const chRes = await fetch(
                `https://graph.microsoft.com/v1.0/teams/${body.teamId}/channels`,
                { headers: { Authorization: `Bearer ${token}` } }
            );
            if (!chRes.ok) throw new Error(`Failed to list channels: ${chRes.status} ${await chRes.text()}`);
            const chData = await chRes.json();
            const channels: any[] = Array.isArray(chData?.value) ? chData.value : [];
            const general =
                channels.find((c) => c.displayName === "General") ??
                channels.find((c) => c.membershipType === "standard");
            if (!general?.id) throw new Error("No standard channel found in this team");
            conversationId = general.id;
        }

        const results = await greetAndRegisterRoster(
            conversationId!,
            process.env.TENANT_ID ?? "",
            { skipExisting: body.skipExisting ?? true, teamId: body.teamId }
        );

        res.writeHead(200);
        res.end(JSON.stringify({
            message: "Roster greeted",
            conversationId,
            summary: {
                total: results.length,
                greeted: results.filter((r) => r.status === "greeted").length,
                alreadyRegistered: results.filter((r) => r.status === "already-registered").length,
                failed: results.filter((r) => r.status === "failed").length,
            },
            results,
        }, null, 2));
    } catch (err: any) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
    }
}

// ── Plain Node handler: GET /api/teamsbot-channel-diagnostics?teamId=... ───────
// Answers the two questions behind BotNotInConversationRoster, using only
// TeamsAppInstallation.ReadWriteForTeam.All + Group.ReadWrite.All (no AppCatalog.Read.All):
//   1. Is our app actually installed in this team?
//   2. Is the target channel standard / private / shared?
export async function handleChannelDiagnostics(req: http.IncomingMessage, res: http.ServerResponse) {
    res.setHeader("Content-Type", "application/json");

    const parsed = new URL(req.url ?? "/", "http://localhost");
    const teamId = parsed.searchParams.get("teamId");
    if (!teamId) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "teamId query parameter is required" }));
        return;
    }

    const externalId = process.env.TEAMS_APP_ID ?? "";

    try {
        const token = await getGraphToken();
        const headers = { Authorization: `Bearer ${token}` };

        const [appsRes, chansRes] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/installedApps?$expand=teamsApp`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels`, { headers }),
        ]);

        const appsRaw = await appsRes.json().catch(() => ({}));
        const chansRaw = await chansRes.json().catch(() => ({}));

        const appName = process.env.TEAMS_APP_NAME ?? "ESI Boomi Integration Agent";
        const installed = Array.isArray(appsRaw?.value) ? appsRaw.value : [];
        // Match on externalId when Graph provides it, else fall back to display name —
        // first-party apps come back with externalId: null, and custom apps sometimes do too.
        const mine = installed.find(
            (a: any) =>
                (externalId && a?.teamsApp?.externalId === externalId) ||
                a?.teamsApp?.displayName === appName
        );

        res.writeHead(200);
        res.end(JSON.stringify({
            teamId,
            lookingForExternalId: externalId,
            appInstalledInTeam: Boolean(mine),
            // The catalog id you can put in TEAMS_CATALOG_APP_ID once known:
            catalogAppId: mine?.teamsApp?.id ?? null,
            matchedByDisplayName: appName,
            installedAppsError: appsRes.ok ? null : { status: appsRes.status, body: appsRaw },
            // Non-Microsoft apps only — your app will be in here once installed.
            customApps: installed
                .filter((a: any) => a?.teamsApp?.distributionMethod !== "store")
                .map((a: any) => ({
                    displayName: a?.teamsApp?.displayName,
                    externalId: a?.teamsApp?.externalId,
                    catalogId: a?.teamsApp?.id,
                    distributionMethod: a?.teamsApp?.distributionMethod,
                })),
            installedApps: installed.map((a: any) => ({
                displayName: a?.teamsApp?.displayName,
                externalId: a?.teamsApp?.externalId,
                catalogId: a?.teamsApp?.id,
            })),
            channelsError: chansRes.ok ? null : { status: chansRes.status, body: chansRaw },
            channels: (Array.isArray(chansRaw?.value) ? chansRaw.value : []).map((c: any) => ({
                channelId: c.id,
                displayName: c.displayName,
                membershipType: c.membershipType, // standard | private | shared
            })),
        }, null, 2));
    } catch (err: any) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
    }
}

// ── Plain Node handler: POST /api/teamsbot-dynamic-channel-creation ────
export async function handleCreateTeamAndChannel(req: http.IncomingMessage, res: http.ServerResponse) {
    res.setHeader("Content-Type", "application/json");

    let body: CreateTeamChannelRequest;
    try {
        body = await readJsonBody(req);
    } catch {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "Invalid JSON body" }));
        return;
    }

    const {
        teamDisplayName,
        teamDescription,
        teamMembers,
        channelDisplayName,
        channelDescription,
        membershipType,
        channelMembers,
    } = body;

    if (!teamDisplayName || !teamMembers?.length) {
        res.writeHead(400);
        res.end(JSON.stringify({
            error: "teamDisplayName and teamMembers (at least one owner required) are required",
        }));
        return;
    }
    if (!channelDisplayName) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "channelDisplayName is required" }));
        return;
    }

    const resolvedMembershipType = membershipType ?? "private";
    if (resolvedMembershipType === "private" && !channelMembers?.length) {
        res.writeHead(400);
        res.end(JSON.stringify({
            error: "channelMembers must be non-empty when membershipType is 'private'",
        }));
        return;
    }

    try {
        const { teamId, failedMembers: failedTeamMembers } = await createTeam(
            teamDisplayName,
            teamDescription ?? "",
            teamMembers
        );

        let result = await createDynamicChannelWithRetry(
            teamId,
            channelDisplayName,
            channelDescription ?? "",
            resolvedMembershipType,
            channelMembers
        );

        // Shared channels provision asynchronously (202) — wait for the channel to exist so we
        // can return its id and install the bot, same as the synchronous (201) path.
        if (result.status === 202 && result.operationLocation) {
            const channel = await resolveAsyncChannel(teamId, result.operationLocation);
            if (channel?.id) result = { status: 201, channel };
        }

        // Install the bot app into the new team + register its channel conversation.
        // Surfaced in the response (not swallowed) — without this the bot is absent from
        // the team roster and every later /api/alert/channel call fails.
        let botInstalled = false;
        let botInstallError: string | null = null;
        if (result.status === 201 && result.channel?.id) {
            try {
                await ensureAppInstalledInTeam(teamId);
                await ensureChannelConversationRef(
                    teamId,
                    result.channel.id,
                    process.env.TENANT_ID
                );
                botInstalled = true;
                console.log(
                    `[Teams] Bot installed + conversation registered for channel ${result.channel.id}`
                );
            } catch (err: any) {
                botInstallError = err.message;
                console.error(
                    `[Teams] Failed to install bot / register conversation: ${err.message}`
                );
            }
        }

        const failedChannelMembers: string[] = [];
        if (result.status === 201 && resolvedMembershipType === "shared" && channelMembers && channelMembers.length > 1) {
            const owner = channelMembers.find((m) => m.role === "owner");
            const rest = channelMembers.filter((m) => m.userId !== owner?.userId);
            await Promise.all(
                rest.map(async (m) => {
                    try {
                        await addChannelMember(teamId, result.channel.id, m);
                    } catch (err: any) {
                        console.error(`Failed to add channel member ${m.userId}: ${err.message}`);
                        failedChannelMembers.push(m.userId);
                    }
                })
            );
        }

        res.writeHead(200);
        res.end(JSON.stringify({
            success: true,
            teamId,
            channelId: result.channel?.id,
            teamName: teamDisplayName,
            channelName: channelDisplayName,
            membershipType: resolvedMembershipType,
            teamMembers: teamMembers.length,
            channelMembers: channelMembers?.length ?? 0,
            failedTeamMembers,
            failedChannelMembers,
            botInstalled,
            botInstallError,
        }));
    } catch (err: any) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
    }
}