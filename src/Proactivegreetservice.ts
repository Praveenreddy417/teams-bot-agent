// ── Proactive greeting / conversation registration ──────────────────────────────
// Removes the "user must message the bot first" requirement for /api/alert/members.
//
// When the bot's app is installed into a team, Teams gives us the team roster. For each
// member we open a 1:1 conversation via the Bot Framework Connector, send the greeting
// card, and persist the ConvRef — so member alerts work immediately afterwards.

import * as https from "https";
import { ClientSecretCredential } from "@azure/identity";
import { conversationStore, ConvRef } from "./Conversationstore";
import { buildGreetingCard } from "./Cardbuilder";

const BOT_APP_ID = process.env.CLIENT_ID ?? "";
const DEFAULT_SERVICE_URL = process.env.BOT_SERVICE_URL ?? "https://smba.trafficmanager.net/in/";

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

export async function getBotFrameworkToken(): Promise<string> {
    const cred = new ClientSecretCredential(
        process.env.TENANT_ID!,
        process.env.CLIENT_ID!,
        process.env.CLIENT_SECRET!
    );
    const resp = await cred.getToken("https://api.botframework.com/.default");
    if (!resp?.token) throw new Error("Empty Bot Framework token");
    return resp.token;
}

export interface RosterMember {
    id: string;                 // Teams user id, e.g. 29:1abc...
    name?: string;
    email?: string;
    userPrincipalName?: string;
    aadObjectId?: string;
    tenantId?: string;
}

/** Read the member roster of a team / channel conversation (Bot Framework, no Graph). */
export async function getConversationMembers(
    conversationId: string,
    token: string,
    serviceUrl = DEFAULT_SERVICE_URL
): Promise<RosterMember[]> {
    const resp = await httpsJson<RosterMember[]>(
        "GET",
        `${serviceUrl}v3/conversations/${encodeURIComponent(conversationId)}/members`,
        token
    );
    if (resp.status < 200 || resp.status >= 300) {
        throw new Error(`Failed to read roster for ${conversationId}: ${resp.status} ${JSON.stringify(resp.data)}`);
    }
    return Array.isArray(resp.data) ? resp.data : [];
}

/** Open (or reuse) the 1:1 chat between the bot and a user; returns the conversation id. */
export async function createPersonalConversation(
    member: RosterMember,
    tenantId: string,
    token: string,
    serviceUrl = DEFAULT_SERVICE_URL
): Promise<string> {
    const resp = await httpsJson<{ id?: string }>(
        "POST",
        `${serviceUrl}v3/conversations`,
        token,
        {
            bot: { id: BOT_APP_ID },
            members: [{ id: member.id }],
            channelData: { tenant: { id: tenantId } },
            isGroup: false,
        }
    );
    if (resp.status < 200 || resp.status >= 300 || !resp.data?.id) {
        throw new Error(`Failed to open 1:1 with ${member.name ?? member.id}: ${resp.status} ${JSON.stringify(resp.data)}`);
    }
    return resp.data.id;
}

function sendCard(conversationId: string, card: object, token: string, serviceUrl = DEFAULT_SERVICE_URL) {
    const activity = {
        type: "message",
        attachments: [
            { contentType: "application/vnd.microsoft.card.adaptive", content: card },
        ],
    };
    return httpsJson("POST", `${serviceUrl}v3/conversations/${encodeURIComponent(conversationId)}/activities`, token, activity);
}

export interface GreetResult {
    user: string;
    email: string;
    status: "greeted" | "already-registered" | "failed";
    error?: string;
}

/**
 * Greet + register every member of a team/channel conversation.
 * `skipExisting` avoids re-greeting users who already have a stored conversation.
 */
export async function greetAndRegisterRoster(
    teamConversationId: string,
    tenantId: string,
    opts: { serviceUrl?: string; skipExisting?: boolean; teamId?: string } = {}
): Promise<GreetResult[]> {
    const serviceUrl = opts.serviceUrl ?? DEFAULT_SERVICE_URL;
    const skipExisting = opts.skipExisting ?? true;

    const token = await getBotFrameworkToken();
    const members = await getConversationMembers(teamConversationId, token, serviceUrl);
    console.log(`[Greet] ${members.length} member(s) in ${teamConversationId}`);

    const results: GreetResult[] = [];

    await Promise.allSettled(
        members.map(async (m) => {
            const email = m.email ?? m.userPrincipalName ?? "";
            const label = m.name ?? m.id;

            // Bots and app-only accounts have no email — skip them.
            if (!email) {
                results.push({ user: label, email: "", status: "failed", error: "Member has no email (likely a bot)" });
                return;
            }

            if (skipExisting && conversationStore.getByEmail(email)) {
                results.push({ user: label, email, status: "already-registered" });
                return;
            }

            try {
                const conversationId = await createPersonalConversation(m, tenantId, token, serviceUrl);

                const ref: ConvRef = {
                    serviceUrl,
                    conversationId,
                    tenantId,
                    botId: BOT_APP_ID,
                    userId: m.id,
                    userName: m.name ?? "",
                    userEmail: email,
                    conversationType: "personal",
                };
                conversationStore.save(m.id, ref);

                // await sendCard(conversationId, buildGreetingCard(m.name ?? "", email), token, serviceUrl);

                console.log(`[Greet] ✅ ${email}`);
                results.push({ user: label, email, status: "greeted" });
            } catch (err: any) {
                console.error(`[Greet] ❌ ${email}: ${err.message}`);
                results.push({ user: label, email, status: "failed", error: err.message });
            }
        })
    );

    return results;
}
