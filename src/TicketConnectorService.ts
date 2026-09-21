// ── Ticket search & update ─────────────────────────────────────────────────────
// Fired by the "/ticket" command flow: search a ticket by incident number,
// then optionally add a work-note update if it's still Open.

import axios from "axios";
import * as https from "https";
import * as dotenv from "dotenv";
dotenv.config({ path: "./env/.env.local" });

// apibaseqa.easystepin.com serves an incomplete TLS chain (missing intermediate) —
// Node rejects it while Postman/browsers tolerate it. Scoped to this host only.
const ticketApiAgent = new https.Agent({ rejectUnauthorized: false });

const TICKET_SEARCH_URL =
    process.env.TICKET_SEARCH_URL ??
    "https://apibaseqa.easystepin.com:9093/ws/rest/teams_bot/Tickets_Connectors/Get_ServiceNow";

const TICKET_UPDATE_URL =
    process.env.TICKET_UPDATE_URL ??
    "https://apibaseqa.easystepin.com:9093/ws/rest/teams_bot/Tickets_Connectors/Update";

const TICKET_CLOSE_URL =
    process.env.TICKET_CLOSE_URL ??
    "https://apibaseqa.easystepin.com:9093/ws/rest/teams_bot/Tickets_Connectors/Close";

// ServiceNow's "Resolved/Closed" state code — fixed for this connector.
const CLOSE_STATE = "7";

const TICKET_AUTH_HEADERS = process.env.BOOMI_AUTH
    ? { Authorization: process.env.BOOMI_AUTH }
    : {};

export type TicketingTool = "ServiceNow" | "ADO";

export interface TicketSearchPayload {
    tickting_tool: TicketingTool;
    incident_nb: string;
}

export interface TicketDetails {
    incident_number: string;
    short_description: string;
    description: string;
    state: string;
    priority: string;
    impact: string;
    urgency: string;
    category: string;
    opened_at: string;
    status: string;
}

export interface TicketSearchResult {
    success: boolean;
    message: string;
    ticket?: TicketDetails;
}

export interface TicketUpdatePayload {
    tickting_tool: TicketingTool;
    incident_nb: string;
    work_notes: string;
}

export interface TicketUpdateResult {
    success: boolean;
    message: string;
    incidentNumber?: string;
}

export interface TicketClosePayload {
    tickting_tool: TicketingTool;
    incident_nb: string;
    close_code: string;
    close_notes: string;
    user_id: string;
}

export interface TicketCloseResult {
    success: boolean;
    message: string;
    incidentNumber?: string;
}

function isSuccess(statusCode: string, statusResp: string): boolean {
    return statusCode === "200" || statusResp.toLowerCase() === "success";
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

function parseSearchResponse(data: any): TicketSearchResult {
    const statusCode = String(data?.Status_Code ?? data?.status_code ?? "");
    const statusResp = String(data?.Status_Response ?? data?.status_response ?? "");
    const success = isSuccess(statusCode, statusResp);

    if (!success) {
        return {
            success: false,
            message: data?.message ?? `No ticket found for that incident number.`,
        };
    }

    return {
        success: true,
        message: "Ticket found.",
        ticket: {
            incident_number: String(data.incident_number ?? ""),
            short_description: String(data.short_description ?? ""),
            description: String(data.description ?? ""),
            state: String(data.state ?? ""),
            priority: String(data.priority ?? ""),
            impact: String(data.impact ?? ""),
            urgency: String(data.urgency ?? ""),
            category: String(data.category ?? ""),
            opened_at: String(data.opened_at ?? ""),
            status: String(data.status ?? ""),
        },
    };
}

function parseUpdateResponse(data: any): TicketUpdateResult {
    const statusCode = String(data?.Status_Code ?? data?.status_code ?? "");
    const statusResp = String(data?.Status_Response ?? data?.status_response ?? "");
    const success = isSuccess(statusCode, statusResp);
    const message =
        data?.Status_Message ?? data?.status_message ?? (success ? "Ticket updated successfully." : "Ticket update failed.");

    return {
        success,
        message: String(message),
        incidentNumber: data?.incident_number ? String(data.incident_number) : undefined,
    };
}

/** POST the search request for the ticket by incident number. Never throws — failures come back as { success: false }. */
export async function searchTicket(payload: TicketSearchPayload): Promise<TicketSearchResult> {
    try {
        console.log("[TicketSearch] POST", TICKET_SEARCH_URL, JSON.stringify(payload));

        const resp = await axios.post(TICKET_SEARCH_URL, payload, {
            headers: { "Content-Type": "application/json", ...TICKET_AUTH_HEADERS },
            timeout: 0, // no timeout — connector can be slow
            httpsAgent: ticketApiAgent,
        });

        const result = parseSearchResponse(parseBody(resp.data));
        console.log(`[TicketSearch] ${result.success ? "✅" : "❌"} ${result.message}`);
        return result;
    } catch (err: any) {
        const raw = err.response?.data;
        if (raw) {
            const parsed = parseSearchResponse(parseBody(raw));
            console.error(`[TicketSearch] ❌ ${err.response.status}: ${parsed.message}`);
            return { ...parsed, success: false };
        }
        console.error("[TicketSearch] ❌", err.message);
        return { success: false, message: err.message ?? "Ticket search request failed." };
    }
}

/** PUT the close request. Never throws — failures come back as { success: false }. */
export async function closeTicket(payload: TicketClosePayload): Promise<TicketCloseResult> {
    const body = { tickting_tool: payload.tickting_tool, incident_nb: payload.incident_nb, state: CLOSE_STATE, close_code: payload.close_code, close_notes: payload.close_notes, user_id: payload.user_id };

    try {
        console.log("[TicketClose] PUT", TICKET_CLOSE_URL, JSON.stringify(body));

        const resp = await axios.put(TICKET_CLOSE_URL, body, {
            headers: { "Content-Type": "application/json", ...TICKET_AUTH_HEADERS },
            timeout: 0, // no timeout — connector can be slow
            httpsAgent: ticketApiAgent,
        });

        const result = parseUpdateResponse(parseBody(resp.data));
        console.log(`[TicketClose] ${result.success ? "✅" : "❌"} ${result.message}`);
        return result;
    } catch (err: any) {
        const raw = err.response?.data;
        if (raw) {
            const parsed = parseUpdateResponse(parseBody(raw));
            console.error(`[TicketClose] ❌ ${err.response.status}: ${parsed.message}`);
            return { ...parsed, success: false };
        }
        console.error("[TicketClose] ❌", err.message);
        return { success: false, message: err.message ?? "Ticket close request failed." };
    }
}

/** POST the work-note update. Never throws — failures come back as { success: false }. */
export async function updateTicket(payload: TicketUpdatePayload): Promise<TicketUpdateResult> {
    try {
        console.log("[TicketUpdate] POST", TICKET_UPDATE_URL, JSON.stringify(payload));

        const resp = await axios.post(TICKET_UPDATE_URL, payload, {
            headers: { "Content-Type": "application/json", ...TICKET_AUTH_HEADERS },
            timeout: 0, // no timeout — connector can be slow
            httpsAgent: ticketApiAgent,
        });

        const result = parseUpdateResponse(parseBody(resp.data));
        console.log(`[TicketUpdate] ${result.success ? "✅" : "❌"} ${result.message}`);
        return result;
    } catch (err: any) {
        const raw = err.response?.data;
        if (raw) {
            const parsed = parseUpdateResponse(parseBody(raw));
            console.error(`[TicketUpdate] ❌ ${err.response.status}: ${parsed.message}`);
            return { ...parsed, success: false };
        }
        console.error("[TicketUpdate] ❌", err.message);
        return { success: false, message: err.message ?? "Ticket update request failed." };
    }
}
