import { buildProcessAlertCard } from "./ProcessAlertCard";

// ── Multi-type alert card builder ───────────────────────────────────────────────
// Dispatches on `type_alert` to render a distinct Adaptive Card per source system:
//   MDM            → Boomi MDM quarantine alert
//   ATOM           → Boomi Atom runtime status
//   SERVICE_NOW    → ServiceNow incident
//   ADO            → Azure DevOps work item
//   PROCESS        → Boomi process execution alert
//   ERROR_HANDLING → Boomi process exception (Log_Data.ErrorRecordDetails)

export type AlertType = "MDM" | "ATOM" | "SERVICE_NOW" | "ADO" | "PROCESS" | "ERROR_HANDLING";

export interface AlertPayload {
    type_alert?: string;
    [key: string]: any;
}

export type CardColor = "default" | "accent" | "good" | "warning" | "attention";

// Fields that must be present per alert type (targetEmails / teamId / channelId are
// validated separately by the endpoint, not here). Dot-separated entries reach into
// nested objects (e.g. ERROR_HANDLING's fields live under Log_Data).
export const REQUIRED_FIELDS_BY_TYPE: Record<AlertType, string[]> = {
    MDM: ["Repository_Name", "Model_Name", "Quarantine_Count"],
    ATOM: ["id", "name", "status"],
    SERVICE_NOW: ["incident_number", "short_description"],
    ADO: ["work_item_id", "title"],
    PROCESS: ["Process_Name", "Process_Id", "Atom_Name", "Execution_Type", "Execution_Duration"],
    ERROR_HANDLING: [
        "Log_Data.Execution_Details.Process_Name",
        "Log_Data.Execution_Details.Execution_Id",
        "Log_Data.ErrorRecordDetails.FailureReason",
    ],
};

// This bot doesn't handle MDM quarantine alerts (that's Data Hub Agent's domain) —
// requests resolving to it (explicit or via missing/unrecognized type_alert defaulting
// to it) are rejected by the endpoint, not silently rendered.
export const DISALLOWED_ALERT_TYPES: AlertType[] = ["MDM"];

export function isAlertTypeAllowed(type: AlertType): boolean {
    return !DISALLOWED_ALERT_TYPES.includes(type);
}

export function resolveAlertType(raw: unknown): AlertType {
    const t = String(raw ?? "MDM").toUpperCase().replace(/[\s\-&]+/g, "_");
    if (t === "MDM" || t === "ATOM" || t === "SERVICE_NOW" || t === "ADO" || t === "PROCESS") return t;
    if (t === "ERROR_EXCEPTING_HANDLING" || t === "ERROR_EXCEPTION_HANDLING" || t === "ERROR_HANDLING") return "ERROR_HANDLING";
    return "MDM";
}

// ── Shared scaffold ──────────────────────────────────────────────────────────────
interface Fact { title: string; value: string; }

interface ScaffoldOpts {
    title: string;
    badge: { label: string; color: CardColor; caption: string };
    sectionTitle: string;
    facts: Fact[];
    note?: string;               // optional longer free-text (e.g. description), rendered wrapped
    highlight?: { text: string; color: CardColor };
    actions?: object[];
    sentBy?: string;
}

export function scaffold(opts: ScaffoldOpts): object {
    const timestamp = new Date().toLocaleString("en-IN", {
        timeZone: "Asia/Kolkata",
        dateStyle: "medium",
        timeStyle: "short",
    });

    const body: object[] = [
        {
            type: "Container",
            style: "emphasis",
            bleed: true,
            items: [
                {
                    type: "ColumnSet",
                    columns: [
                        {
                            type: "Column",
                            width: "stretch",
                            items: [
                                { type: "TextBlock", text: opts.title, weight: "Bolder", size: "Large", wrap: true },
                                {
                                    type: "TextBlock",
                                    text: `${timestamp} IST${opts.sentBy ? `  ·  triggered by ${opts.sentBy}` : ""}`,
                                    size: "Small",
                                    isSubtle: true,
                                    spacing: "None",
                                    wrap: true,
                                },
                            ],
                        },
                        {
                            type: "Column",
                            width: "auto",
                            verticalContentAlignment: "Center",
                            items: [
                                {
                                    type: "TextBlock",
                                    text: opts.badge.label,
                                    weight: "Bolder",
                                    color: opts.badge.color,
                                    size: "Medium",
                                    horizontalAlignment: "Right",
                                    wrap: true,
                                },
                                {
                                    type: "TextBlock",
                                    text: opts.badge.caption,
                                    size: "Small",
                                    isSubtle: true,
                                    horizontalAlignment: "Right",
                                    spacing: "None",
                                },
                            ],
                        },
                    ],
                },
            ],
        },
        {
            type: "Container",
            spacing: "Medium",
            items: [
                { type: "TextBlock", text: opts.sectionTitle, weight: "Bolder", size: "Medium" },
                {
                    type: "FactSet",
                    facts: opts.facts
                        .filter((f) => f.value !== undefined && f.value !== null && f.value !== "")
                        .map((f) => ({ title: f.title, value: f.value })),
                },
                ...(opts.note
                    ? [{
                        type: "TextBlock",
                        text: opts.note,
                        wrap: true,
                        isSubtle: true,
                        spacing: "Small",
                    }]
                    : []),
            ],
        },
    ];

    if (opts.highlight) {
        body.push({
            type: "Container",
            spacing: "Small",
            style: "emphasis",
            items: [
                { type: "TextBlock", text: opts.highlight.text, weight: "Bolder", color: opts.highlight.color, wrap: true },
            ],
        });
    }

    return {
        type: "message",
        attachments: [
            {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: {
                    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                    type: "AdaptiveCard",
                    version: "1.5",
                    body,
                    ...(opts.actions?.length ? { actions: opts.actions } : {}),
                },
            },
        ],
    };
}

export const str = (v: unknown, fallback = "—"): string =>
    v === undefined || v === null || v === "" ? fallback : String(v);

// ── MDM ──────────────────────────────────────────────────────────────────────────
function buildMdmCard(d: AlertPayload, sentBy?: string): object {
    const count = parseInt(String(d.Quarantine_Count ?? "0"), 10);
    const badge =
        count > 50 ? { label: "Critical", color: "attention" as CardColor } :
            count > 20 ? { label: "High", color: "warning" as CardColor } :
                count > 5 ? { label: "Medium", color: "accent" as CardColor } :
                    { label: "Normal", color: "good" as CardColor };

    return scaffold({
        title: "⚠️  MDM Quarantine Alert",
        badge: { ...badge, caption: "SEVERITY" },
        sectionTitle: "Quarantine Details",
        facts: [
            { title: "Type", value: str(d.type_alert, "MDM") },
            { title: "Repository", value: str(d.Repository_Name) },
            { title: "Model", value: str(d.Model_Name) },
            { title: "Quarantine Count", value: str(d.Quarantine_Count, "0") },
        ],
        highlight: { text: `🗂  Quarantined records: ${str(d.Quarantine_Count, "0")}`, color: badge.color },
        actions: [
            {
                type: "Action.Submit",
                title: "🎫 Create Ticket",
                data: {
                    actionType: "createTicket",
                    alertType: "MDM",
                    Repository_Name: d.Repository_Name,
                    Model_Name: d.Model_Name,
                    Quarantine_Count: d.Quarantine_Count,
                    severity: badge.label,
                },
            },
        ],
        sentBy,
    });
}

// ── Boomi Atom ─────────────────────────────────────────────────────────────────────
function buildAtomCard(d: AlertPayload, sentBy?: string): object {
    const status = String(d.status ?? "").toUpperCase();
    const badge =
        status === "ONLINE" ? { label: str(d.status, "Online"), color: "good" as CardColor } :
            status === "OFFLINE" ? { label: str(d.status, "Offline"), color: "attention" as CardColor } :
                { label: str(d.status, "Unknown"), color: "warning" as CardColor };

    return scaffold({
        title: "🔷  Boomi Atom Status",
        badge: { ...badge, caption: "STATUS" },
        sectionTitle: "Atom Details",
        facts: [
            { title: "Name", value: str(d.name) },
            { title: "Status", value: str(d.status) },
            { title: "Type", value: str(d.type) },
            { title: "Host Name", value: str(d.hostName) },
            { title: "Current Version", value: str(d.currentVersion) },
            { title: "Date Installed", value: str(d.dateInstalled) },
            { title: "Purge History (days)", value: str(d.purgeHistoryDays) },
            { title: "Atom ID", value: str(d.id) },
            { title: "Created By", value: str(d.createdBy) },
        ],
        highlight: { text: `🖥  ${str(d.name)} is ${str(d.status)}`, color: badge.color },
        sentBy,
    });
}

// ── ServiceNow incident ──────────────────────────────────────────────────────────
function buildServiceNowCard(d: AlertPayload, sentBy?: string): object {
    const prio = parseInt(String(d.priority ?? "").match(/\d+/)?.[0] ?? "0", 10);
    const badge =
        prio === 1 ? { label: str(d.priority, "P1"), color: "attention" as CardColor } :
            prio === 2 ? { label: str(d.priority, "P2"), color: "warning" as CardColor } :
                prio === 3 ? { label: str(d.priority, "P3"), color: "accent" as CardColor } :
                    { label: str(d.priority, "Low"), color: "good" as CardColor };

    const actions: object[] = [];
    if (d.url) {
        actions.push({ type: "Action.OpenUrl", title: "🔗 Open Incident", url: String(d.url) });
    }

    return scaffold({
        title: "🎫  ServiceNow Incident",
        badge: { ...badge, caption: "PRIORITY" },
        sectionTitle: "Incident Details",
        facts: [
            { title: "Incident", value: str(d.incident_number) },
            { title: "State", value: str(d.state) },
            { title: "Priority", value: str(d.priority) },
            { title: "Impact", value: str(d.impact) },
            { title: "Urgency", value: str(d.urgency) },
            { title: "Category", value: str(d.category) },
            { title: "Opened At", value: str(d.opened_at) },
            { title: "Status", value: str(d.status) },
        ],
        note: d.description ? str(d.description) : undefined,
        highlight: { text: `📝  ${str(d.short_description)}`, color: badge.color },
        actions,
        sentBy,
    });
}

// ── Azure DevOps work item ───────────────────────────────────────────────────────
function buildAdoCard(d: AlertPayload, sentBy?: string): object {
    const state = String(d.state ?? "").toUpperCase();
    const badge =
        ["DONE", "CLOSED", "RESOLVED", "COMPLETED"].includes(state) ? { label: str(d.state, "Done"), color: "good" as CardColor } :
            ["DOING", "ACTIVE", "IN PROGRESS"].includes(state) ? { label: str(d.state, "Active"), color: "warning" as CardColor } :
                { label: str(d.state, "To Do"), color: "accent" as CardColor };

    const actions: object[] = [];
    if (d.url) {
        actions.push({ type: "Action.OpenUrl", title: "🔗 Open Work Item", url: String(d.url) });
    }

    return scaffold({
        title: "🔵  Azure DevOps Work Item",
        badge: { ...badge, caption: "STATE" },
        sectionTitle: "Work Item Details",
        facts: [
            { title: "ID", value: str(d.work_item_id) },
            { title: "Type", value: str(d.type) },
            { title: "State", value: str(d.state) },
            { title: "Priority", value: str(d.priority) },
            { title: "Project", value: str(d.project) },
            { title: "Created By", value: str(d.created_by) },
            { title: "Created Date", value: str(d.created_date) },
        ],
        note: d.description ? str(d.description) : undefined,
        highlight: { text: `🧩  ${str(d.title)}`, color: badge.color },
        actions,
        sentBy,
    });
}

// ── Boomi process exception ───────────────────────────────────────────────────────
function buildErrorHandlingCard(d: AlertPayload, sentBy?: string): object {
    const execution = d.Log_Data?.Execution_Details ?? {};
    const errorDetails = d.Log_Data?.ErrorRecordDetails ?? {};

    return scaffold({
        title: "❌  Process Exception",
        badge: { label: "Error", color: "attention", caption: "STATUS" },
        sectionTitle: "Execution Details",
        facts: [
            { title: "Process Name", value: str(execution.Process_Name) },
            { title: "Execution ID", value: str(execution.Execution_Id) },
        ],
        note: str(errorDetails.FailureReason, ""),
        highlight: { text: `🚨  ${str(execution.Process_Name)} failed`, color: "attention" },
        sentBy,
    });
}

// ── Dispatcher ───────────────────────────────────────────────────────────────────
export function buildAlertCard(body: AlertPayload, sentBy?: string): object {
    switch (resolveAlertType(body.type_alert)) {
        case "ATOM": return buildAtomCard(body, sentBy);
        case "SERVICE_NOW": return buildServiceNowCard(body, sentBy);
        case "ADO": return buildAdoCard(body, sentBy);
        case "PROCESS": return buildProcessAlertCard(body, sentBy);
        case "ERROR_HANDLING": return buildErrorHandlingCard(body, sentBy);
        case "MDM":
        default: return buildMdmCard(body, sentBy);
    }
}
