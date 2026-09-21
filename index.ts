import * as dotenv from "dotenv";
dotenv.config({ path: "./env/.env.local" });

import * as http from "http";
import * as https from "https";
import app from "./app";
import { conversationStore, ConvRef } from "./src/Conversationstore";
import {
  handleListTeams,
  handleCreateTeamAndChannel,
  handleChannelDiagnostics,
  handleGreetTeam,
} from "./src/Teamschannelservice";
import {
  handleListChannelMembers,
  handleUpdateChannel,
  handleDeleteChannel,
  handleDeleteTeam,
} from "./src/Channelmanageservice";
import { postCardToChannel } from "./src/Channelinstallservice";
import {
  buildAlertCard,
  resolveAlertType,
  isAlertTypeAllowed,
  REQUIRED_FIELDS_BY_TYPE,
  AlertPayload,
} from "./src/AlertCards";

// Resolve a dot-separated path (e.g. "Log_Data.ErrorRecordDetails.FailureReason") against a payload.
function getByPath(obj: any, path: string): any {
  return path.split(".").reduce((o, k) => (o == null ? undefined : o[k]), obj);
}

// Validate required fields for whichever alert type the payload declares.
function missingFieldsForType(body: AlertPayload): string[] {
  const type = resolveAlertType(body.type_alert);
  return REQUIRED_FIELDS_BY_TYPE[type].filter((f) => {
    const v = getByPath(body, f);
    return v === undefined || v === null || v === "";
  });
}

// Render (and most PaaS hosts) expose exactly one port externally, so the bot
// messaging endpoint and the Alert API must share it. The bot's App listens
// on an internal-only port; the external listener proxies /api/messages to
// it and handles every other /api/* route itself.
const EXTERNAL_PORT = parseInt(process.env.PORT ?? "3978", 10);
const BOT_INTERNAL_PORT = parseInt(process.env.BOT_INTERNAL_PORT ?? String(EXTERNAL_PORT + 1), 10);
const MESSAGING_ENDPOINT = "/api/messages";

// ── Token acquisition ─────────────────────────────────────────────────────────
async function getBotToken(): Promise<string> {
  const { ClientSecretCredential, ManagedIdentityCredential } = await import("@azure/identity");

  console.log("[Token] CLIENT_ID:", process.env.CLIENT_ID ? "✅" : "❌ missing");
  console.log(process.env.CLIENT_ID, "process.env.CLIENT_ID")
  console.log("[Token] CLIENT_SECRET:", process.env.CLIENT_SECRET ? "✅" : "❌ missing");
  console.log("[Token] TENANT_ID:", process.env.TENANT_ID ? "✅" : "❌ missing");

  if (process.env.CLIENT_ID && process.env.CLIENT_SECRET && process.env.TENANT_ID) {
    console.log("[Token] → ClientSecretCredential (local dev)");
    const cred = new ClientSecretCredential(
      process.env.TENANT_ID,
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET
    );
    const resp = await cred.getToken("https://api.botframework.com/.default");
    if (!resp?.token) throw new Error("Empty token from ClientSecretCredential");
    return resp.token;
  }

  console.log("[Token] → ManagedIdentityCredential (Azure)");
  const cred = new ManagedIdentityCredential({ clientId: process.env.CLIENT_ID });
  const resp = await cred.getToken("https://api.botframework.com/.default");
  if (!resp?.token) throw new Error("Empty token from ManagedIdentityCredential");
  return resp.token;
}

// ── Send proactive card to a Teams conversation ────────────────────────────────
function sendActivity(ref: ConvRef, cardJson: string, token: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const serviceUrl = ref.serviceUrl ?? "https://smba.trafficmanager.net/in/";
    const url = new URL(`/v3/conversations/${ref.conversationId}/activities`, serviceUrl);
    const protocol = url.protocol === "https:" ? https : http;
    const bodyBuf = Buffer.from(cardJson, "utf-8");

    const req = protocol.request(
      {
        hostname: url.hostname,
        port: url.port || (url.protocol === "https:" ? 443 : 80),
        path: url.pathname,
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${token}`,
          "Content-Length": bodyBuf.byteLength,
        },
      },
      (res) => {
        let data = "";
        res.on("data", (c) => (data += c));
        res.on("end", () => {
          if (res.statusCode && res.statusCode >= 200 && res.statusCode < 300) resolve();
          else reject(new Error(`Bot Framework ${res.statusCode}: ${data}`));
        });
      }
    );
    req.on("error", reject);
    req.write(bodyBuf);
    req.end();
  });
}

// ── Helper: read HTTP request body ────────────────────────────────────────────
function readBody(req: http.IncomingMessage): Promise<unknown> {
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

// ── Alert / Channel-Creation API handler ────────────────────────────────────────
async function apiHandler(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
  const url = (req.url ?? "/").split("?")[0];
  const method = req.method ?? "GET";

  res.setHeader("Content-Type", "application/json");
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-api-key");

  if (method === "OPTIONS") { res.writeHead(204); res.end(); return; }

  // ── Dynamic team + channel creation (raw Node handlers, no Express) ─
  if (method === "GET" && url === "/api/teamsbot-list-teams") {
    await handleListTeams(req, res);
    return;
  }
  if (method === "POST" && url === "/api/teamsbot-dynamic-channel-creation") {
    await handleCreateTeamAndChannel(req, res);
    return;
  }
  if (method === "GET" && url === "/api/teamsbot-channel-diagnostics") {
    await handleChannelDiagnostics(req, res);
    return;
  }
  if (method === "POST" && url === "/api/teamsbot-greet-team") {
    await handleGreetTeam(req, res);
    return;
  }

  // ── Channel management (update / delete) ────────────────────────
  if (method === "GET" && url === "/api/teamsbot-channel-members") {
    await handleListChannelMembers(req, res);
    return;
  }
  if (method === "POST" && url === "/api/teamsbot-channel-update") {
    await handleUpdateChannel(req, res);
    return;
  }
  if ((method === "POST" || method === "DELETE") && url === "/api/teamsbot-channel-delete") {
    await handleDeleteChannel(req, res);
    return;
  }
  if ((method === "POST" || method === "DELETE") && url === "/api/teamsbot-team-delete") {
    await handleDeleteTeam(req, res);
    return;
  }

  // ── Health ──────────────────────────────────────────────────────
  if (method === "GET" && url === "/api/health") {
    res.writeHead(200);
    res.end(JSON.stringify({
      status: "ok",
      service: "integration-agent-bot",
      timestamp: new Date().toISOString(),
      activeConversations: conversationStore.size(),
      breakdown: {
        personal: conversationStore.getAllPersonal().length,
        group: conversationStore.getAllGroups().filter((r: any) => r.conversationType === "groupChat").length,
        channel: conversationStore.getAllGroups().filter((r: any) => r.conversationType === "channel").length,
      },
    }));
    return;
  }

  // ── Integration alert broadcast to ALL registered conversations ──
  if (method === "POST" && url === "/api/alert") {

    const apiKey = req.headers["x-api-key"];
    if (process.env.ALERT_API_KEY && apiKey !== process.env.ALERT_API_KEY) {
      res.writeHead(401);
      res.end(JSON.stringify({ error: "Unauthorized: invalid x-api-key" }));
      return;
    }

    let body: AlertPayload;
    try {
      body = (await readBody(req)) as AlertPayload;
    } catch {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "Invalid JSON body" }));
      return;
    }

    if (!isAlertTypeAllowed(resolveAlertType(body.type_alert))) {
      res.writeHead(400);
      res.end(JSON.stringify({
        error: "Alert type not supported by this bot",
        type_alert: body.type_alert,
        hint: "MDM alerts are handled by Data Hub Agent, not Team's Bot Boomi Agent.",
      }));
      return;
    }

    const missing = missingFieldsForType(body);
    if (missing.length) {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "Missing required fields", type_alert: resolveAlertType(body.type_alert), missing }));
      return;
    }

    if (conversationStore.size() === 0) {
      res.writeHead(202);
      res.end(JSON.stringify({
        message: "No conversations registered yet.",
        hint: "Open the bot in Teams and send any message first.",
      }));
      return;
    }

    const card = buildAlertCard(body);
    const cardStr = JSON.stringify(card);
    const results: { target: string; type: string; status: string; error?: string }[] = [];

    let token: string;
    try {
      token = await getBotToken();
    } catch (err: any) {
      console.error("[Alert] Token error:", err.message);
      res.writeHead(500);
      res.end(JSON.stringify({ error: "Failed to acquire bot token", detail: err.message }));
      return;
    }

    const allTargets = conversationStore.getAll();
    console.log(`[Alert] Sending to ${allTargets.length} conversation(s)...`);

    await Promise.allSettled(
      allTargets.map(async (ref: any) => {
        const label = ref.conversationType === "personal"
          ? (ref.userEmail || ref.userName)
          : ref.conversationId;
        try {
          await sendActivity(ref, cardStr, token);
          console.log(`[Alert] ✅ ${label} (${ref.conversationType})`);
          results.push({ target: label, type: ref.conversationType, status: "sent" });
        } catch (err: any) {
          console.error(`[Alert] ❌ ${label}:`, err.message);
          results.push({ target: label, type: ref.conversationType, status: "failed", error: err.message });
        }
      })
    );

    res.writeHead(200);
    res.end(JSON.stringify({
      message: "Alert dispatched",
      timestamp: new Date().toISOString(),
      summary: {
        total: allTargets.length,
        sent: results.filter(r => r.status === "sent").length,
        failed: results.filter(r => r.status === "failed").length,
        personal: results.filter(r => r.type === "personal").length,
        group: results.filter(r => r.type === "groupChat").length,
        channel: results.filter(r => r.type === "channel").length,
      },
      results,
    }));
    return;
  }

  // ── Integration alert to SPECIFIC member emails only ─────────────────────
  if (method === "POST" && url === "/api/alert/members") {

    const apiKey = req.headers["x-api-key"];
    if (process.env.ALERT_API_KEY && apiKey !== process.env.ALERT_API_KEY) {
      res.writeHead(401);
      res.end(JSON.stringify({ error: "Unauthorized: invalid x-api-key" }));
      return;
    }

    let body: AlertPayload & { targetEmails?: string[] };
    try {
      body = (await readBody(req)) as any;
    } catch {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "Invalid JSON body" }));
      return;
    }

    if (!body.targetEmails?.length) {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "targetEmails (non-empty array of emails) is required" }));
      return;
    }

    if (!isAlertTypeAllowed(resolveAlertType(body.type_alert))) {
      res.writeHead(400);
      res.end(JSON.stringify({
        error: "Alert type not supported by this bot",
        type_alert: body.type_alert,
        hint: "MDM alerts are handled by Data Hub Agent, not Team's Bot Boomi Agent.",
      }));
      return;
    }

    const missing = missingFieldsForType(body);
    if (missing.length) {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "Missing required fields", type_alert: resolveAlertType(body.type_alert), missing }));
      return;
    }

    const card = buildAlertCard(body);
    const cardStr = JSON.stringify(card);

    let token: string;
    try {
      token = await getBotToken();
    } catch (err: any) {
      console.error("[Alert] Token error:", err.message);
      res.writeHead(500);
      res.end(JSON.stringify({ error: "Failed to acquire bot token", detail: err.message }));
      return;
    }

    const results: { target: string; status: string; error?: string }[] = [];

    await Promise.allSettled(
      body.targetEmails!.map(async (email) => {
        const ref = conversationStore.getByEmail(email);
        if (!ref) {
          results.push({
            target: email,
            status: "failed",
            error: "No registered conversation for this email — user must message the bot first",
          });
          return;
        }
        try {
          await sendActivity(ref, cardStr, token);
          console.log(`[Alert] ✅ ${email} (personal)`);
          results.push({ target: email, status: "sent" });
        } catch (err: any) {
          console.error(`[Alert] ❌ ${email}:`, err.message);
          results.push({ target: email, status: "failed", error: err.message });
        }
      })
    );

    res.writeHead(200);
    res.end(JSON.stringify({
      message: "Alert dispatched to specific members",
      timestamp: new Date().toISOString(),
      summary: {
        total: body.targetEmails!.length,
        sent: results.filter((r) => r.status === "sent").length,
        failed: results.filter((r) => r.status === "failed").length,
      },
      results,
    }));
    return;
  }

  // ── Integration alert to a SPECIFIC team's channel only ───────────────────
  if (method === "POST" && url === "/api/alert/channel") {

    const apiKey = req.headers["x-api-key"];
    if (process.env.ALERT_API_KEY && apiKey !== process.env.ALERT_API_KEY) {
      res.writeHead(401);
      res.end(JSON.stringify({ error: "Unauthorized: invalid x-api-key" }));
      return;
    }

    let body: AlertPayload & {
      channelId?: string | string[];
      teamId?: string | string[];
    };
    try {
      body = (await readBody(req)) as any;
    } catch {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "Invalid JSON body" }));
      return;
    }

    // Accept both the legacy single-string shape and the newer array shape.
    const teamIds = Array.isArray(body.teamId) ? body.teamId : body.teamId ? [body.teamId] : [];
    const channelIds = Array.isArray(body.channelId) ? body.channelId : body.channelId ? [body.channelId] : [];

    if (!teamIds.length || !channelIds.length) {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "channelId and teamId are both required" }));
      return;
    }
    if (teamIds.length !== channelIds.length) {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "teamId and channelId arrays must be the same length (paired by index)" }));
      return;
    }

    if (!isAlertTypeAllowed(resolveAlertType(body.type_alert))) {
      res.writeHead(400);
      res.end(JSON.stringify({
        error: "Alert type not supported by this bot",
        type_alert: body.type_alert,
        hint: "MDM alerts are handled by Data Hub Agent, not Team's Bot Boomi Agent.",
      }));
      return;
    }

    const missing = missingFieldsForType(body);
    if (missing.length) {
      res.writeHead(400);
      res.end(JSON.stringify({ error: "Missing required fields", type_alert: resolveAlertType(body.type_alert), missing }));
      return;
    }

    const card = buildAlertCard(body);
    const cardStr = JSON.stringify(card);

    const results: { teamId: string; channelId: string; status: string; strategy?: string; error?: string }[] = [];

    await Promise.allSettled(
      teamIds.map(async (teamId, i) => {
        const channelId = channelIds[i];
        try {
          const { strategy } = await postCardToChannel(teamId, channelId, cardStr, process.env.TENANT_ID);
          console.log(`[Alert] ✅ channel ${channelId} (${strategy})`);
          results.push({ teamId, channelId, status: "sent", strategy });
        } catch (err: any) {
          console.error(`[Alert] ❌ channel ${channelId}:`, err.message);
          results.push({ teamId, channelId, status: "failed", error: err.message });
        }
      })
    );

    const allFailed = results.every((r) => r.status === "failed");
    res.writeHead(allFailed ? 500 : 200);
    res.end(JSON.stringify({
      message: "Alert dispatched to channel(s)",
      timestamp: new Date().toISOString(),
      summary: {
        total: results.length,
        sent: results.filter((r) => r.status === "sent").length,
        failed: results.filter((r) => r.status === "failed").length,
      },
      results,
    }));
    return;
  }

  res.writeHead(404);
  res.end(JSON.stringify({
    error: "Not found",
    received: { method, url },
    available: [
      "GET /api/health",
      "POST /api/alert",
      "POST /api/alert/members",
      "POST /api/alert/channel",
      "GET /api/teamsbot-list-teams",
      "POST /api/teamsbot-dynamic-channel-creation",
      "GET /api/teamsbot-channel-diagnostics?teamId=...",
      "POST /api/teamsbot-greet-team",
    ],
  }));
}

// ── Proxy Teams messaging traffic to the bot's internal-only port ──────────────
function proxyToBot(req: http.IncomingMessage, res: http.ServerResponse): void {
  const proxyReq = http.request(
    {
      hostname: "127.0.0.1",
      port: BOT_INTERNAL_PORT,
      path: req.url,
      method: req.method,
      headers: req.headers,
    },
    (proxyRes) => {
      res.writeHead(proxyRes.statusCode ?? 502, proxyRes.headers);
      proxyRes.pipe(res);
    }
  );
  proxyReq.on("error", (err) => {
    console.error("❌ Bot proxy error:", err.message);
    if (!res.headersSent) res.writeHead(502, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "Bot server unreachable" }));
  });
  req.pipe(proxyReq);
}

// ── Single externally-reachable server: routes bot traffic vs. Alert API ───────
const externalServer = http.createServer((req, res) => {
  const url = (req.url ?? "/").split("?")[0];
  if (url === MESSAGING_ENDPOINT || url.startsWith(`${MESSAGING_ENDPOINT}/`)) {
    proxyToBot(req, res);
    return;
  }
  apiHandler(req, res);
});

// ── Crash guards ───────────────────────────────────────────────────────────────
process.on("uncaughtException", (err) => console.error("💥 UNCAUGHT:", err));
process.on("unhandledRejection", (err) => console.error("💥 UNHANDLED:", err));

externalServer.on("error", (err: NodeJS.ErrnoException) => {
  if (err.code === "EADDRINUSE")
    console.error(`❌ Port ${EXTERNAL_PORT} already in use.`);
  else
    console.error("❌ External server error:", err.message);
  process.exit(1);
});

// ── Start both servers ────────────────────────────────────────────────────────
(async () => {
  try {
    await app.start(BOT_INTERNAL_PORT);
    console.log(`\n✅ Bot started (internal) → http://127.0.0.1:${BOT_INTERNAL_PORT}`);
  } catch (err: any) {
    console.error("❌ Bot failed to start:", err.message);
    process.exit(1);
  }

  externalServer.listen(EXTERNAL_PORT, "0.0.0.0", () => {
    console.log(`💬 Messaging endpoint     → http://localhost:${EXTERNAL_PORT}${MESSAGING_ENDPOINT}`);
    console.log(`📡 Alert API (all)        → http://localhost:${EXTERNAL_PORT}/api/alert`);
    console.log(`📧 Alert API (members)    → http://localhost:${EXTERNAL_PORT}/api/alert/members`);
    console.log(`💬 Alert API (channel)    → http://localhost:${EXTERNAL_PORT}/api/alert/channel`);
    console.log(`❤️  Health                 → http://localhost:${EXTERNAL_PORT}/api/health`);
    console.log(`🔗 List teams              → http://localhost:${EXTERNAL_PORT}/api/teamsbot-list-teams`);
    console.log(`🔗 Create channel          → http://localhost:${EXTERNAL_PORT}/api/teamsbot-dynamic-channel-creation`);
    console.log(`\n🚀 Integration Agent bot running!\n`);
  });
})();