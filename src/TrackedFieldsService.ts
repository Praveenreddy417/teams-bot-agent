// ── TrackedFields lookup ────────────────────────────────────────────────────────
// Fired by the "/track" command: query a Boomi tracked field by process name,
// tracking field value, and connector side. Response shape varies per process —
// callers map whatever keys come back rather than assuming a fixed schema.

import axios from "axios";
import * as https from "https";
import * as dotenv from "dotenv";
dotenv.config({ path: "./env/.env.local" });

// apibaseqa.easystepin.com serves an incomplete TLS chain (missing intermediate) —
// Node rejects it while Postman/browsers tolerate it. Scoped to this host only.
const trackedFieldsApiAgent = new https.Agent({ rejectUnauthorized: false });

const TRACKED_FIELDS_URL =
    process.env.TRACKED_FIELDS_URL ??
    "https://apibaseqa.easystepin.com:9093/ws/rest/teams_bot/TrackedFields/Query";

const TRACKED_FIELDS_AUTH_HEADERS = process.env.BOOMI_AUTH
    ? { Authorization: process.env.BOOMI_AUTH }
    : {};

export type ConnectorSide = "Source" | "Target";

export interface TrackedFieldsPayload {
    processname: string;
    trackingField: string;
    userselectconnector: ConnectorSide;
}

export interface TrackedFieldsResult {
    success: boolean;
    message: string;
    data?: Record<string, any>;
}

function parseBody(raw: any): any {
    if (typeof raw !== "string") return raw;
    const trimmed = raw.trim();
    if (!trimmed) return {};
    try {
        return JSON.parse(trimmed);
    } catch {
        return { message: raw };
    }
}

/** POST the tracked-field lookup. Never throws — failures come back as { success: false }. */
export async function queryTrackedField(payload: TrackedFieldsPayload): Promise<TrackedFieldsResult> {
    try {
        console.log("[TrackedFields] POST", TRACKED_FIELDS_URL, JSON.stringify(payload));

        const resp = await axios.post(TRACKED_FIELDS_URL, payload, {
            headers: { "Content-Type": "application/json", ...TRACKED_FIELDS_AUTH_HEADERS },
            timeout: 0, // no timeout — connector can be slow
            httpsAgent: trackedFieldsApiAgent,
        });

        const data = parseBody(resp.data);
        if (data && typeof data === "object" && !Array.isArray(data) && Object.keys(data).length) {
            console.log(`[TrackedFields] ✅ ${Object.keys(data).length} field(s) returned`);
            return { success: true, message: "Record found.", data };
        }
        console.log("[TrackedFields] ❌ No matching record.");
        return { success: false, message: "No record found for that tracking field." };
    } catch (err: any) {
        const raw = err.response?.data;
        if (raw) {
            const parsed = parseBody(raw);
            const message = typeof parsed === "string" ? parsed : parsed?.message ?? `Request failed (${err.response.status}).`;
            console.error(`[TrackedFields] ❌ ${err.response.status}: ${message}`);
            return { success: false, message };
        }
        console.error("[TrackedFields] ❌", err.message);
        return { success: false, message: err.message ?? "Tracked field request failed." };
    }
}
