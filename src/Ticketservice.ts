// ── Ticket creation ─────────────────────────────────────────────────────────────
// Fired by the "Create Ticket" Action.Submit on the MDM alert card.
// Turns the alert's fields into a ticket payload and posts it to the Tickets connector.

import axios from "axios";
import * as https from "https";
import * as dotenv from "dotenv";
dotenv.config({ path: "./env/.env.local" });

// apibaseqa.easystepin.com serves an incomplete TLS chain (missing intermediate) —
// Node rejects it while Postman/browsers tolerate it. Scoped to this host only.
const ticketApiAgent = new https.Agent({ rejectUnauthorized: false });

const TICKET_AUTH_HEADERS = process.env.BOOMI_AUTH
    ? { Authorization: process.env.BOOMI_AUTH }
    : {};

const TICKET_CREATE_URL =
    process.env.TICKET_CREATE_URL ??
    "https://apibaseqa.easystepin.com:9093/ws/rest/teams_bot/Tickets_Connectors/create";

// Fixed for this project — the connector only serves the datahub custom project at P1.
const PROJECT_TYPE = "integrations custom";
const PRIORITY = "1";

export interface TicketPayload {
    project_type: string;
    title: string;
    description: string;
    priority: string;
    user_id: string;
}

export interface TicketResult {
    success: boolean;
    message: string;
    ticketId?: string;
    raw?: any;
}

/** Values carried on the card's Action.Submit `data` object. */
export interface MdmTicketSource {
    Repository_Name?: string;
    Model_Name?: string;
    Quarantine_Count?: string | number;
    severity?: string;
}

const dash = (v: unknown, fallback = "Unknown") =>
    v === undefined || v === null || v === "" ? fallback : String(v);

/** "MDM Quarantine — CustomerMDM / Customer (12 records)" */
export function buildTicketTitle(src: MdmTicketSource): string {
    const repo = dash(src.Repository_Name);
    const model = dash(src.Model_Name);
    return `MDM Quarantine Alert — ${repo} ${model}`;
}

/** One readable sentence built from the alert fields. */
export function buildTicketDescription(src: MdmTicketSource): string {
    const repo = dash(src.Repository_Name);
    const model = dash(src.Model_Name);
    const count = dash(src.Quarantine_Count, "0");
    const plural = count === "1" ? "record has" : "records have";

    let text =
        `${count} ${plural} been quarantined in the ${repo} repository for the ${model} model. ` +
        `Please review the quarantined records and resolve the underlying data quality issues.`;

    if (src.severity) text += ` Severity reported by the alert: ${src.severity}.`;
    return text;
}

// Fallback for malformed JSON (the connector sometimes omits a comma when
// incident_number is empty). Pulls known fields out with regex, comma or not.
function extractFieldsLoosely(raw: string): any {
    const field = (key: string) => raw.match(new RegExp(`"${key}"\\s*:\\s*"([^"]*)"`))?.[1];
    return {
        Status_Code: field("Status_Code"),
        Status_Response: field("Status_Response"),
        Status_Message: field("Status_Message"),
        incident_number: field("incident_number"),
    };
}

function parseTicketBody(raw: any): any {
    if (typeof raw !== "string" || !raw.trim()) return raw;
    try {
        return JSON.parse(raw);
    } catch {
        return extractFieldsLoosely(raw);
    }
}

function parseTicketResponse(data: any): TicketResult {
    if (data === undefined || data === null) {
        return { success: true, message: "Ticket created." };
    }

    const statusCode = String(data.Status_Code ?? "");
    const statusResp = String(data.Status_Response ?? "");
    const message = data.Status_Message;
    const ticketId = data.incident_number;

    const success = statusCode === "200" && statusResp.toLowerCase() === "success";

    return {
        success,
        message: message ? String(message) : success ? "Ticket created." : "Ticket creation failed.",
        ticketId: ticketId ? String(ticketId) : undefined,
        raw: data,
    };
}

/** POST the ticket. Never throws — failures come back as { success: false }. */
export async function createTicket(payload: TicketPayload): Promise<TicketResult> {
    try {
        console.log("[Ticket] POST", TICKET_CREATE_URL, JSON.stringify(payload));

        const resp = await axios.post(TICKET_CREATE_URL, payload, {
            headers: { "Content-Type": "application/json", ...TICKET_AUTH_HEADERS },
            timeout: 0, // no timeout — connector can be slow
            httpsAgent: ticketApiAgent,
        });

        const result = parseTicketResponse(parseTicketBody(resp.data));
        console.log(`[Ticket] ${result.success ? "✅" : "❌"} ${result.message}`);
        return result;
    } catch (err: any) {
        const raw = err.response?.data;
        if (raw) {
            const parsed = parseTicketResponse(parseTicketBody(raw));
            // An HTTP error status means failure regardless of the body's wording.
            console.error(`[Ticket] ❌ ${err.response.status}: ${parsed.message}`);
            return { ...parsed, success: false };
        }
        console.error("[Ticket] ❌", err.message);
        return { success: false, message: err.message ?? "Ticket request failed." };
    }
}

/** Build the payload from an MDM alert and create the ticket. */
export function createMdmTicket(src: MdmTicketSource, userId: string): Promise<TicketResult> {
    return createTicket({
        project_type: PROJECT_TYPE,
        title: buildTicketTitle(src),
        description: buildTicketDescription(src),
        priority: PRIORITY,
        user_id: userId,
    });
}

/** Values carried on the Process alert card's Action.Submit `data` object. */
export interface ProcessTicketSource {
    Process_Name?: string;
    Process_Id?: string;
    Atom_Name?: string;
    Execution_Type?: string;
    Execution_Duration?: string | number;
}

/** "Process Alert — testing_Delay_Process (Atom 004)" */
export function buildProcessTicketTitle(src: ProcessTicketSource): string {
    const process = dash(src.Process_Name);
    const atom = dash(src.Atom_Name);
    return `Process Alert — ${process} (${atom})`;
}

/** One readable sentence built from the process alert fields. */
export function buildProcessTicketDescription(src: ProcessTicketSource): string {
    const process = dash(src.Process_Name);
    const processId = dash(src.Process_Id);
    const atom = dash(src.Atom_Name);
    const execType = dash(src.Execution_Type);
    const durationMs = parseInt(String(src.Execution_Duration ?? ""), 10);
    const duration = Number.isFinite(durationMs) && durationMs > 0
        ? `${(durationMs / 1000 / 60).toFixed(1)} minutes`
        : dash(src.Execution_Duration, "an unknown duration");

    return (
        `Process ${process} (ID: ${processId}) ran on ${atom} via ${execType} execution and took ${duration}. ` +
        `Please review the process run and confirm whether follow-up action is needed.`
    );
}

/** Build the payload from a Process alert and create the ticket. */
export function createProcessTicket(src: ProcessTicketSource, userId: string): Promise<TicketResult> {
    return createTicket({
        project_type: PROJECT_TYPE,
        title: buildProcessTicketTitle(src),
        description: buildProcessTicketDescription(src),
        priority: PRIORITY,
        user_id: userId,
    });
}
