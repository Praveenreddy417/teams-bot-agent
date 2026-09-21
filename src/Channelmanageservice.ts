// ── Channel management: update (name/description + members/roles) and delete ────
// Self-contained — does not touch the existing creation flow in Teamschannelservice.ts.
//
// NOTE (Teams platform rule): member management works only on `private` and `shared`
// channels. A `standard` channel inherits its membership from the team, so Graph rejects
// per-channel member add/remove/role changes there. Handlers below fail fast with a clear
// message instead of surfacing a raw Graph error.

import * as http from "http";
import * as dotenv from "dotenv";
dotenv.config({ path: "./env/.env.local" });

// ── Graph app-only token (own cache, independent of other services) ────────────
let cachedToken: { token: string; expiresAt: number } | null = null;

async function getGraphToken(): Promise<string> {
    if (cachedToken && cachedToken.expiresAt > Date.now() + 60_000) return cachedToken.token;

    const tenantId = process.env.AAD_TENANT_ID ?? process.env.TENANT_ID!;
    const clientId = process.env.AAD_CLIENT_ID ?? process.env.CLIENT_ID!;
    const clientSecret = process.env.AAD_CLIENT_SECRET ?? process.env.CLIENT_SECRET!;

    const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            scope: "https://graph.microsoft.com/.default",
            grant_type: "client_credentials",
        }).toString(),
    });
    if (!res.ok) throw new Error(`Graph token fetch failed: ${res.status} ${await res.text()}`);

    const data = await res.json();
    cachedToken = { token: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
    return cachedToken.token;
}

/** Pull the human-readable message out of a Graph error body, so callers get a plain string. */
function graphErrorMessage(text: string): string {
    if (!text) return "";
    try {
        const parsed = JSON.parse(text);
        return parsed?.error?.message ?? parsed?.message ?? text;
    } catch {
        return text;
    }
}

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

export interface ChannelMemberChange {
    userId: string;                  // AAD object id OR userPrincipalName/email
    role?: "owner" | "member";
}

interface GraphChannel {
    id: string;
    displayName: string;
    description?: string;
    membershipType?: "standard" | "private" | "shared";
}

interface GraphChannelMember {
    id: string;                      // membership id (needed for PATCH/DELETE)
    displayName?: string;
    userId?: string;                 // AAD object id
    email?: string;
    roles?: string[];
}

// ── Graph helpers ──────────────────────────────────────────────────────────────
async function getChannel(teamId: string, channelId: string, token: string): Promise<GraphChannel> {
    const res = await fetch(
        `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}`,
        { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`Channel lookup failed: ${res.status} ${await res.text()}`);
    return res.json();
}

/**
 * Where membership actually lives for this channel type:
 *   private / shared → the channel's own roster
 *   standard         → the team's roster (standard channels inherit team membership,
 *                      Graph rejects per-channel member calls on them)
 */
function membersBaseUrl(teamId: string, channelId: string, membershipType?: string): string {
    return membershipType === "standard"
        ? `https://graph.microsoft.com/v1.0/teams/${teamId}/members`
        : `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/members`;
}

async function listMembersAt(baseUrl: string, token: string): Promise<GraphChannelMember[]> {
    const res = await fetch(baseUrl, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`Member list failed: ${res.status} ${await res.text()}`);
    const data = await res.json();
    return Array.isArray(data?.value) ? data.value : [];
}

async function listMembers(teamId: string, channelId: string, token: string): Promise<GraphChannelMember[]> {
    return listMembersAt(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/members`, token);
}

/** Match a member by AAD object id or email — callers may pass either. */
function findMember(members: GraphChannelMember[], userId: string): GraphChannelMember | undefined {
    const needle = userId.toLowerCase();
    return members.find(
        (m) => m.userId?.toLowerCase() === needle || m.email?.toLowerCase() === needle
    );
}


// ── GET /api/teamsbot-channel-members?teamId=...&channelId=... ─────────────────
export async function handleListChannelMembers(req: http.IncomingMessage, res: http.ServerResponse) {
    res.setHeader("Content-Type", "application/json");

    const parsed = new URL(req.url ?? "/", "http://localhost");
    const teamId = parsed.searchParams.get("teamId");
    const channelId = parsed.searchParams.get("channelId");

    if (!teamId || !channelId) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "teamId and channelId query parameters are required" }));
        return;
    }

    try {
        const token = await getGraphToken();
        const [channel, members] = await Promise.all([
            getChannel(teamId, channelId, token),
            listMembers(teamId, channelId, token),
        ]);

        res.writeHead(200);
        res.end(JSON.stringify({
            teamId,
            channelId,
            displayName: channel.displayName,
            membershipType: channel.membershipType,
            memberManagementSupported: channel.membershipType !== "standard",
            members: members.map((m) => ({
                membershipId: m.id,
                displayName: m.displayName,
                userId: m.userId,
                email: m.email,
                roles: m.roles ?? [],
            })),
        }, null, 2));
    } catch (err: any) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
    }
}

// ── POST /api/teamsbot-channel-update ──────────────────────────────────────────
// Body: { teamId, channelId, displayName?, description?,
//         addMembers?: [{userId, role}], removeMembers?: ["id-or-email"],
//         updateMembers?: [{userId, role}] }
export async function handleUpdateChannel(req: http.IncomingMessage, res: http.ServerResponse) {
    res.setHeader("Content-Type", "application/json");

    let body: {
        teamId?: string;
        channelId?: string;
        displayName?: string;
        description?: string;
        teamDisplayName?: string;
        teamDescription?: string;
        membershipType?: string;
        addMembers?: ChannelMemberChange[];
        removeMembers?: string[];
        updateMembers?: ChannelMemberChange[];
    };
    try {
        body = await readJsonBody(req);
    } catch {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "Invalid JSON body" }));
        return;
    }

    const { teamId, channelId } = body;
    if (!teamId) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "teamId is required" }));
        return;
    }

    // Graph accepts membershipType only at creation — it is immutable afterwards.
    if (body.membershipType !== undefined) {
        res.writeHead(400);
        res.end(JSON.stringify({
            error: "membershipType cannot be changed after a channel is created",
            detail:
                "Microsoft Graph only accepts membershipType at creation time. To switch a channel " +
                "between standard/private/shared you must create a new channel with the desired type " +
                "(messages and files do not carry over).",
        }));
        return;
    }

    const wantsMemberChanges =
        Boolean(body.addMembers?.length || body.removeMembers?.length || body.updateMembers?.length);
    const wantsChannelEdit = body.displayName !== undefined || body.description !== undefined;
    const wantsTeamEdit = body.teamDisplayName !== undefined || body.teamDescription !== undefined;

    if (!wantsMemberChanges && !wantsChannelEdit && !wantsTeamEdit) {
        res.writeHead(400);
        res.end(JSON.stringify({
            error:
                "Nothing to update — provide teamDisplayName/teamDescription, displayName/description, " +
                "and/or addMembers/removeMembers/updateMembers",
        }));
        return;
    }

    // channelId is only needed for channel-scoped work; a pure team rename doesn't require it.
    if ((wantsMemberChanges || wantsChannelEdit) && !channelId) {
        res.writeHead(400);
        res.end(JSON.stringify({
            error: "channelId is required when updating channel properties or members",
        }));
        return;
    }

    try {
        const token = await getGraphToken();
        const channel = channelId ? await getChannel(teamId, channelId, token) : undefined;

        const result: any = {
            teamId,
            channelId,
            membershipType: channel?.membershipType,
            teamUpdated: false,
            teamUpdateError: null as string | null,
            channelUpdated: false,
            channelUpdateError: null as string | null,
            added: [] as any[],
            removed: [] as any[],
            roleUpdated: [] as any[],
        };

        // ── 0. Team name / description ──
        // Try the Teams resource first, then fall back to the backing group (covered by
        // Group.ReadWrite.All). The team name mirrors the group's displayName.
        if (wantsTeamEdit) {
            const teamPatch: Record<string, string> = {};
            if (body.teamDisplayName !== undefined) teamPatch.displayName = body.teamDisplayName;
            if (body.teamDescription !== undefined) teamPatch.description = body.teamDescription;

            const tryPatch = async (url: string) => {
                const r = await fetch(url, {
                    method: "PATCH",
                    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
                    body: JSON.stringify(teamPatch),
                });
                return { ok: r.ok, status: r.status, text: r.ok ? "" : await r.text() };
            };

            const viaTeam = await tryPatch(`https://graph.microsoft.com/v1.0/teams/${teamId}`);
            if (viaTeam.ok) {
                result.teamUpdated = true;
                result.teamUpdateVia = "teams";
            } else {
                const viaGroup = await tryPatch(`https://graph.microsoft.com/v1.0/groups/${teamId}`);
                if (viaGroup.ok) {
                    result.teamUpdated = true;
                    result.teamUpdateVia = "groups";
                } else {
                    result.teamUpdateError =
                        `Team update failed (${viaTeam.status}): ${graphErrorMessage(viaTeam.text)} | ` +
                        `fallback via groups (${viaGroup.status}): ${graphErrorMessage(viaGroup.text)}`;
                }
            }
        }

        // ── 1. Channel properties (works on every channel type) ──
        // Non-fatal: a rename can fail independently of the member changes below, so record
        // the error and keep going rather than aborting the whole request.
        if (wantsChannelEdit) {
            const patchChannel = async (patchBody: Record<string, string>) => {
                const patchRes = await fetch(
                    `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}`,
                    {
                        method: "PATCH",
                        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
                        body: JSON.stringify(patchBody),
                    }
                );
                return { ok: patchRes.ok, status: patchRes.status, text: patchRes.ok ? "" : await patchRes.text() };
            };

            const full: Record<string, string> = {};
            if (body.displayName !== undefined) full.displayName = body.displayName;
            if (body.description !== undefined) full.description = body.description;

            const first = await patchChannel(full);
            if (first.ok) {
                result.channelUpdated = true;
            } else if (first.status === 412 && full.displayName !== undefined && full.description !== undefined) {
                // Renaming needs the channel's SharePoint folder, which is provisioned lazily on
                // brand-new channels. Description alone doesn't, so salvage that part.
                const retry = await patchChannel({ description: full.description });
                result.channelUpdated = retry.ok;
                result.descriptionUpdated = retry.ok;
                result.displayNameUpdated = false;
                result.channelUpdateError =
                    `Rename skipped (${first.status}): ${graphErrorMessage(first.text)}. ` +
                    `The channel's SharePoint folder does not exist yet — open the channel's Files tab ` +
                    `in Teams once (or add any file), then retry the rename. Description was updated.`;
            } else {
                result.channelUpdated = false;
                result.channelUpdateError =
                    first.status === 412
                        ? `Channel update failed (${first.status}): ${graphErrorMessage(first.text)}. ` +
                          `The channel's SharePoint folder does not exist yet — open the channel's Files tab ` +
                          `in Teams once (or add any file), then retry.`
                        : `Channel update failed (${first.status}): ${graphErrorMessage(first.text)}`;
            }
        }

        // ── 2. Membership changes — works on all three channel types ──
        // Standard channels have no roster of their own, so these are applied to the team.
        // channelId/channel are guaranteed here by the validation above, but narrow explicitly
        // so this type-checks under strictNullChecks.
        if (wantsMemberChanges && channelId && channel) {
            const base = membersBaseUrl(teamId, channelId, channel.membershipType);
            const scope = channel.membershipType === "standard" ? "team" : "channel";
            result.memberScope = scope;
            if (scope === "team") {
                result.note =
                    "This is a standard channel — its membership is inherited from the team, " +
                    "so member changes were applied at the team level.";
            }

            // Add
            for (const m of body.addMembers ?? []) {
                try {
                    const addRes = await fetch(base, {
                        method: "POST",
                        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
                        body: JSON.stringify({
                            "@odata.type": "#microsoft.graph.aadUserConversationMember",
                            roles: m.role === "owner" ? ["owner"] : [],
                            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${m.userId}')`,
                        }),
                    });
                    if (!addRes.ok) throw new Error(`(${addRes.status}) ${graphErrorMessage(await addRes.text())}`);
                    result.added.push({ userId: m.userId, role: m.role ?? "member", scope, status: "added" });
                } catch (err: any) {
                    result.added.push({ userId: m.userId, scope, status: "failed", error: err.message });
                }
            }

            // Remove / role update both need the membership id — read the roster once.
            if (body.removeMembers?.length || body.updateMembers?.length) {
                const members = await listMembersAt(base, token);

                for (const userId of body.removeMembers ?? []) {
                    const match = findMember(members, userId);
                    if (!match) {
                        result.removed.push({ userId, scope, status: "failed", error: `Not a member of this ${scope}` });
                        continue;
                    }
                    try {
                        const delRes = await fetch(`${base}/${match.id}`, {
                            method: "DELETE",
                            headers: { Authorization: `Bearer ${token}` },
                        });
                        if (!delRes.ok) throw new Error(`(${delRes.status}) ${graphErrorMessage(await delRes.text())}`);
                        result.removed.push({ userId, scope, status: "removed" });
                    } catch (err: any) {
                        result.removed.push({ userId, scope, status: "failed", error: err.message });
                    }
                }

                for (const m of body.updateMembers ?? []) {
                    const match = findMember(members, m.userId);
                    if (!match) {
                        result.roleUpdated.push({ userId: m.userId, scope, status: "failed", error: `Not a member of this ${scope}` });
                        continue;
                    }
                    try {
                        const patchRes = await fetch(`${base}/${match.id}`, {
                            method: "PATCH",
                            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
                            body: JSON.stringify({
                                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                                roles: m.role === "owner" ? ["owner"] : [],
                            }),
                        });
                        if (!patchRes.ok) throw new Error(`(${patchRes.status}) ${graphErrorMessage(await patchRes.text())}`);
                        result.roleUpdated.push({ userId: m.userId, role: m.role ?? "member", scope, status: "updated" });
                    } catch (err: any) {
                        result.roleUpdated.push({ userId: m.userId, scope, status: "failed", error: err.message });
                    }
                }
            }
        }

        result.success = true;
        res.writeHead(200);
        res.end(JSON.stringify(result, null, 2));
    } catch (err: any) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
    }
}

// ── POST | DELETE /api/teamsbot-team-delete ────────────────────────────────────
// Body: { teamId, confirm: true }
// Deletes the whole team — every channel, the mailbox, the SharePoint site and its files.
// Requires an explicit `confirm: true` so a stray teamId can't wipe a team by accident.
// Soft delete: recoverable for ~30 days via /directory/deletedItems.
export async function handleDeleteTeam(req: http.IncomingMessage, res: http.ServerResponse) {
    res.setHeader("Content-Type", "application/json");

    let body: { teamId?: string; confirm?: boolean };
    try {
        body = await readJsonBody(req);
    } catch {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "Invalid JSON body" }));
        return;
    }

    const parsed = new URL(req.url ?? "/", "http://localhost");
    const teamId = body.teamId ?? parsed.searchParams.get("teamId") ?? undefined;

    if (!teamId) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "teamId is required" }));
        return;
    }

    if (body.confirm !== true) {
        res.writeHead(400);
        res.end(JSON.stringify({
            error: "Refusing to delete without confirmation",
            detail:
                "Deleting a team removes all its channels, conversations, the group mailbox, and the " +
                "SharePoint site with every file in it. Send \"confirm\": true in the body to proceed.",
        }));
        return;
    }

    try {
        const token = await getGraphToken();

        // Read the name first so the response can state exactly what was deleted.
        let teamName: string | undefined;
        try {
            const lookup = await fetch(
                `https://graph.microsoft.com/v1.0/groups/${teamId}?$select=displayName`,
                { headers: { Authorization: `Bearer ${token}` } }
            );
            if (lookup.ok) teamName = (await lookup.json())?.displayName;
        } catch {
            // Non-fatal — proceed and let the delete call report any real problem.
        }

        const delRes = await fetch(`https://graph.microsoft.com/v1.0/groups/${teamId}`, {
            method: "DELETE",
            headers: { Authorization: `Bearer ${token}` },
        });
        if (!delRes.ok) {
            throw new Error(`Team delete failed (${delRes.status}): ${graphErrorMessage(await delRes.text())}`);
        }

        res.writeHead(200);
        res.end(JSON.stringify({
            success: true,
            message: "Team deleted — recoverable for ~30 days",
            teamId,
            teamName,
            restoreHint: `POST https://graph.microsoft.com/v1.0/directory/deletedItems/${teamId}/restore`,
        }));
    } catch (err: any) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
    }
}

// ── POST | DELETE /api/teamsbot-channel-delete ─────────────────────────────────
// Body: { teamId, channelId }
// Soft delete — Teams keeps the channel recoverable for ~30 days. "General" cannot be deleted.
export async function handleDeleteChannel(req: http.IncomingMessage, res: http.ServerResponse) {
    res.setHeader("Content-Type", "application/json");

    let body: { teamId?: string; channelId?: string };
    try {
        body = await readJsonBody(req);
    } catch {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "Invalid JSON body" }));
        return;
    }

    // Allow query params too, so a bare DELETE without a body works.
    const parsed = new URL(req.url ?? "/", "http://localhost");
    const teamId = body.teamId ?? parsed.searchParams.get("teamId") ?? undefined;
    const channelId = body.channelId ?? parsed.searchParams.get("channelId") ?? undefined;

    if (!teamId || !channelId) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "teamId and channelId are both required" }));
        return;
    }

    try {
        const token = await getGraphToken();

        // Read first so the response can name what was deleted, and so we can block General.
        let channelName: string | undefined;
        try {
            const channel = await getChannel(teamId, channelId, token);
            channelName = channel.displayName;
            if (channel.displayName === "General") {
                res.writeHead(400);
                res.end(JSON.stringify({
                    error: "The General channel cannot be deleted — delete the team instead.",
                }));
                return;
            }
        } catch {
            // Lookup failure shouldn't block the delete attempt; Graph will report the real error.
        }

        const delRes = await fetch(
            `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}`,
            { method: "DELETE", headers: { Authorization: `Bearer ${token}` } }
        );
        if (!delRes.ok) {
            throw new Error(`Channel delete failed: ${delRes.status} ${await delRes.text()}`);
        }

        res.writeHead(200);
        res.end(JSON.stringify({
            success: true,
            message: "Channel deleted (recoverable for ~30 days)",
            teamId,
            channelId,
            channelName,
        }));
    } catch (err: any) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
    }
}
