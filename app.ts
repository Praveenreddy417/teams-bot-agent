// app.ts
import { stripMentionsText, TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { LocalStorage } from "@microsoft/teams.common";
import { ManagedIdentityCredential } from "@azure/identity";
import * as dotenv from "dotenv";

import config from "./config";
import { buildBoomiResponseCard, buildErrorCard, buildGreetingCard, buildHelpCard, buildWhoAmICard } from "./src/Cardbuilder";
import { conversationStore, ConvRef } from "./src/Conversationstore";
import { askChat } from "./src/chatService";
import { buildChatChartCard, buildChatTableCard, buildChatTextCard } from "./src/Chatcardbuilder";
import { greetAndRegisterRoster } from "./src/Proactivegreetservice";
import {
  createMdmTicket,
  buildTicketTitle,
  buildTicketDescription,
  createProcessTicket,
  buildProcessTicketTitle,
  buildProcessTicketDescription,
} from "./src/Ticketservice";
import { buildTicketResultCard } from "./src/Ticketresultcard";
import { searchTicket, updateTicket, closeTicket, TicketingTool } from "./src/TicketConnectorService";
import {
  buildTicketSearchCard,
  buildTicketDetailCard,
  buildUpdateNotesCard,
  buildUpdateResultCard,
  buildCloseTicketCard,
  buildCloseResultCard,
} from "./src/TicketSearchCard";
import { queryTrackedField, ConnectorSide } from "./src/TrackedFieldsService";
import { buildTrackedFieldsInputCard, buildTrackedFieldsResultCard } from "./src/TrackedFieldsCard";
// import { authStore } from "./src/authStore";
// import { requestOtp, validateOtp } from "./src/authService";

dotenv.config({ path: "./env/.env.local" });

// ── Storage ───────────────────────────────────────────────────────────────────
const storage = new LocalStorage();

// ── Token factory (Managed Identity for Azure, skipped for local dev) ─────────
const createTokenFactory = () =>
  async (scope: string | string[], tenantId?: string): Promise<string> => {
    const cred = new ManagedIdentityCredential({ clientId: process.env.CLIENT_ID });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const resp = await cred.getToken(scopes, { tenantId });
    return resp.token;
  };

const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

const credentialOptions =
  config.MicrosoftAppType === "UserAssignedMsi" ? { ...tokenCredentials } : undefined;

// ── App ───────────────────────────────────────────────────────────────────────
const app = new App({ ...credentialOptions, storage });

// ── Session store (userId → Boomi session_id for conversation continuity) ─────
const boomiSessions = new Map<string, string>();

// ── Graph app-only token (cached, same pattern as Teamschannelservice.ts) ─────
let cachedGraphToken: { token: string; expiresAt: number } | null = null;

async function getGraphToken(): Promise<string | null> {
  if (cachedGraphToken && cachedGraphToken.expiresAt > Date.now() + 60_000) {
    return cachedGraphToken.token;
  }

  const tenantId = process.env.AAD_TENANT_ID ?? process.env.TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID ?? process.env.CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET ?? process.env.CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    console.warn("[App] Graph credentials missing — cannot resolve real email, falling back to name.");
    return null;
  }

  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
        }).toString(),
      }
    );
    if (!res.ok) {
      console.error("[App] Graph token fetch failed:", res.status, await res.text());
      return null;
    }
    const data = await res.json();
    cachedGraphToken = { token: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
    return cachedGraphToken.token;
  } catch (err) {
    console.error("[App] Graph token fetch error:", err);
    return null;
  }
}

// ── Resolve a user's real email via Graph, using the AAD object ID Teams gives us ──
async function resolveUserEmail(aadObjectId: string | undefined, fallbackName: string): Promise<string> {
  if (!aadObjectId) return fallbackName;

  const token = await getGraphToken();
  if (!token) return fallbackName;

  try {
    const userRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${aadObjectId}?$select=mail,userPrincipalName`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!userRes.ok) {
      console.warn(`[App] Graph user lookup failed for ${aadObjectId}: ${userRes.status}`);
      return fallbackName;
    }
    const userData = await userRes.json();
    return userData.mail ?? userData.userPrincipalName ?? fallbackName;
  } catch (err) {
    console.error("[App] Failed to resolve email via Graph:", err);
    return fallbackName;
  }
}

// ── Install handler ───────────────────────────────────────────────────────────
// Fires when the app is installed (personal OR team scope). Greeting each member here
// registers their 1:1 conversation, so /api/alert/members works without them having to
// message the bot first.
app.on("install.add", async (context) => {
  try {
    const activity: any = context.activity;
    const conversationType = activity.conversation?.conversationType ?? "personal";
    const tenantId = activity.channelData?.tenant?.id ?? process.env.TENANT_ID ?? "";

    console.log(`[Install] App installed — scope=${conversationType}`);

    // Personal install: this conversation IS the user's 1:1 chat.
    if (conversationType === "personal") {
      const userEmail = await resolveUserEmail(
        activity.from?.aadObjectId,
        activity.from?.name ?? ""
      );
      conversationStore.save(activity.from.id, {
        serviceUrl: activity.serviceUrl,
        conversationId: activity.conversation.id,
        tenantId,
        botId: activity.recipient?.id ?? "",
        userId: activity.from.id,
        userName: activity.from?.name ?? "",
        userEmail,
        conversationType: "personal",
      });
      await context.send(buildGreetingCard(activity.from?.name ?? "", userEmail));
      console.log(`[Install] Greeted ${userEmail}`);
      return;
    }

    // Team install: greet + register every member of the team.
    const results = await greetAndRegisterRoster(activity.conversation.id, tenantId, {
      serviceUrl: activity.serviceUrl,
      teamId: activity.channelData?.team?.id,
    });
    console.log(
      `[Install] Roster greeted — ${results.filter((r) => r.status === "greeted").length} new, ` +
      `${results.filter((r) => r.status === "already-registered").length} existing, ` +
      `${results.filter((r) => r.status === "failed").length} failed`
    );
  } catch (err) {
    console.error("[Install] Failed to greet on install:", err);
  }
});

// ── Message handler ───────────────────────────────────────────────────────────
app.on("message", async (context) => {
  try {
    const activity = context.activity;
    const userId = activity.from.id;

    // Adaptive Card Action.Submit payloads arrive as activity.value (no text).
    if (activity.value) {
      const value: any = activity.value;

      if (value.actionType === "createTicket") {
        const createdBy = await resolveUserEmail(
          (activity.from as any).aadObjectId,
          activity.from.name ?? ""
        );

        const isProcessAlert = value.alertType === "PROCESS";

        const ticketPayload = {
          project_type: "integrations custom",
          title: isProcessAlert ? buildProcessTicketTitle(value) : buildTicketTitle(value),
          description: isProcessAlert ? buildProcessTicketDescription(value) : buildTicketDescription(value),
          priority: "1",
          user_id: createdBy,
        };

        await context.send("⏳ Creating ticket...");

        const result = isProcessAlert ? await createProcessTicket(value, createdBy) : await createMdmTicket(value, createdBy);
        await context.send(buildTicketResultCard(result, ticketPayload, createdBy) as any);
        return;
      }

      if (value.actionType === "searchTicket") {
        const tool = (value.tickting_tool ?? "ServiceNow") as TicketingTool;
        const incidentNb = String(value.incident_nb ?? "").trim();

        if (!incidentNb) {
          await context.send(buildTicketSearchCard("Enter an incident number.") as any);
          return;
        }

        const searchResult = await searchTicket({ tickting_tool: tool, incident_nb: incidentNb });
        if (!searchResult.success || !searchResult.ticket) {
          await context.send(buildTicketSearchCard(searchResult.message) as any);
          return;
        }

        await context.send(buildTicketDetailCard(searchResult.ticket, tool) as any);
        return;
      }

      if (value.actionType === "openUpdateNotes") {
        const tool = (value.tickting_tool ?? "ServiceNow") as TicketingTool;
        await context.send(buildUpdateNotesCard(String(value.incident_nb ?? ""), tool) as any);
        return;
      }

      if (value.actionType === "submitTicketUpdate") {
        const tool = (value.tickting_tool ?? "ServiceNow") as TicketingTool;
        const incidentNb = String(value.incident_nb ?? "");
        const workNotes = String(value.work_notes ?? "");

        await context.send("⏳ Updating ticket...");

        const updateResult = await updateTicket({ tickting_tool: tool, incident_nb: incidentNb, work_notes: workNotes });
        await context.send(buildUpdateResultCard(updateResult, incidentNb) as any);
        return;
      }

      if (value.actionType === "openCloseTicket") {
        const tool = (value.tickting_tool ?? "ServiceNow") as TicketingTool;
        await context.send(buildCloseTicketCard(String(value.incident_nb ?? ""), tool) as any);
        return;
      }

      if (value.actionType === "submitTicketClose") {
        const tool = (value.tickting_tool ?? "ServiceNow") as TicketingTool;
        const incidentNb = String(value.incident_nb ?? "");
        const closeCode = String(value.close_code ?? "");
        const closeNotes = String(value.close_notes ?? "");
        const closedBy = await resolveUserEmail(
          (activity.from as any).aadObjectId,
          activity.from.name ?? ""
        );

        await context.send("⏳ Closing ticket...");

        const closeResult = await closeTicket({ tickting_tool: tool, incident_nb: incidentNb, close_code: closeCode, close_notes: closeNotes, user_id: closedBy });
        await context.send(buildCloseResultCard(closeResult, incidentNb) as any);
        return;
      }

      if (value.actionType === "queryTrackedField") {
        const processname = String(value.processname ?? "").trim();
        const trackingField = String(value.trackingField ?? "").trim();
        const userselectconnector = (value.userselectconnector ?? "Source") as ConnectorSide;

        if (!processname || !trackingField || !userselectconnector) {
          await context.send(buildTrackedFieldsInputCard("Process Name, Tracking Field and Connector are all required.") as any);
          return;
        }

        await context.send("⏳ Looking up tracked field...");

        const trackResult = await queryTrackedField({ processname, trackingField, userselectconnector });
        await context.send(buildTrackedFieldsResultCard(trackResult, { processname, trackingField, userselectconnector }) as any);
        return;
      }

      return;
    }

    const text = (stripMentionsText(activity) ?? "").trim();
    const textLower = text.toLowerCase();

    // Resolve real email via Graph (falls back to display name if Graph lookup fails)
    const userEmail = await resolveUserEmail(
      (activity.from as any).aadObjectId,
      activity.from.name ?? ""
    );

    console.log(`[Bot] "${text}" from ${activity.from?.name} (${activity.conversation.conversationType})`);

    const conversationType = (activity.conversation.conversationType ?? "personal") as
      | "personal"
      | "groupChat"
      | "channel";

    const ref: ConvRef = {
      serviceUrl: activity.serviceUrl,
      conversationId: activity.conversation.id,
      tenantId: (activity.channelData as any)?.tenant?.id ?? "",
      botId: activity.recipient.id,
      userId: activity.from.id,
      userName: activity.from.name ?? "",
      userEmail,
      conversationType,
      channelId: (activity.channelData as any)?.channel?.id,
    };

    if (conversationType === "personal") {
      conversationStore.save(userId, ref);
    } else {
      conversationStore.saveGroup(activity.conversation.id, ref);
    }

    if (["hi", "hello", "hey"].includes(textLower)) {
      await context.send(buildGreetingCard(activity.from.name ?? "", userEmail));
      return;
    }

    if (textLower === "/help") {
      await context.send(buildHelpCard(userEmail));
      return;
    }

    if (textLower === "/whoami") {
      await context.send(
        buildWhoAmICard(
          activity.from.name ?? "",
          userEmail,
          activity.from.id,
          conversationType,
          activity.conversation.id
        )
      );
      return;
    }

    if (textLower === "/ticket") {
      await context.send(buildTicketSearchCard() as any);
      return;
    }

    if (textLower === "/track") {
      await context.send(buildTrackedFieldsInputCard() as any);
      return;
    }

    if (textLower === "/diag") {
      await context.send(JSON.stringify(activity, null, 2).substring(0, 3000));
      return;
    }

    if (!text) return;

    await context.send("⏳ Querying Integration Analytics...");

    const existingSessionId = boomiSessions.get(userId) ?? null;

    let chatResp;
    try {
      chatResp = await askChat(text, existingSessionId, userEmail);
    } catch (err: any) {
      console.error("[Bot] askChat failed:", err.message);
      await context.send(
        buildErrorCard(`Unable to reach the Integration Analytics agent.\n\n_${err.message}_`)
      );
      return;
    }

    if (chatResp.session_id) {
      boomiSessions.set(userId, chatResp.session_id);
      console.log(`[Bot] Stored session_id=${chatResp.session_id} for userId=${userId}`);
    }

    if (chatResp.type === "text" && chatResp.answer) {
      if (chatResp.mode?.toLowerCase().includes("custom")) {
        console.log(`[Bot] mode="${chatResp.mode}" → buildChatTextCard`);
        await context.send(buildChatTextCard(chatResp.answer));
      } else {
        // Missing mode (server doesn't always send it) defaults to markdown rendering —
        // "agent" answers (and most untagged ones) are markdown-formatted.
        console.log(`[Bot] mode="${chatResp.mode}" → buildBoomiResponseCard`);
        await context.send(buildBoomiResponseCard(chatResp.answer, [], text));
      }
      return;
    }

    if (chatResp.type === "table" && chatResp.data) {
      console.log(`[Bot] mode="${chatResp.mode}" → buildChatTableCard`);
      const { title, columns, rows } = chatResp.data;
      await context.send(buildChatTableCard(title ?? "Integration results", columns ?? [], rows ?? []));
      return;
    }

    if (chatResp.type === "chart" && chatResp.data) {
      const { title, chart_type, card, note } = chatResp.data;
      console.log(`[Bot] chart_type="${chart_type}" → buildChatChartCard`);
      if (!card) {
        await context.send(buildChatTextCard("Chart could not be rendered."));
        return;
      }
      await context.send(buildChatChartCard(chart_type, title, card, note));
      return;
    }

  } catch (err) {
    console.error("[Bot] Unhandled error in message handler:", err);
    await context.send(buildErrorCard("An internal error occurred. Please try again."));
  }
});

export default app;