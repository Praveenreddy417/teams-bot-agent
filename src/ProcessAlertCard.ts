// ── Boomi Process execution alert ───────────────────────────────────────────────
// Built from a `type_alert: "PROCESS"` payload (Process_Name, Process_Id, Atom_Name,
// Execution_Type, Execution_Duration). Dispatched via buildAlertCard in AlertCards.ts.

import { scaffold, str, CardColor, AlertPayload } from "./AlertCards";

export interface ProcessAlertPayload extends AlertPayload {
    Process_Name?: string;
    Process_Id?: string;
    Atom_Name?: string;
    Execution_Type?: string;
    Execution_Duration?: string | number;
}

function durationLabel(ms: unknown): string {
    const n = parseInt(String(ms ?? ""), 10);
    if (!Number.isFinite(n) || n <= 0) return str(ms, "0");
    return `${(n / 1000 / 60).toFixed(1)} min`;
}

export function buildProcessAlertCard(d: ProcessAlertPayload, sentBy?: string): object {
    return scaffold({
        title: "⚙️  Process Execution Alert",
        badge: { label: str(d.Execution_Type, "—"), color: "accent" as CardColor, caption: "EXECUTION TYPE" },
        sectionTitle: "Process Details",
        facts: [
            { title: "Process Name", value: str(d.Process_Name) },
            { title: "Process ID", value: str(d.Process_Id) },
            { title: "Atom", value: str(d.Atom_Name) },
            { title: "Execution Type", value: str(d.Execution_Type) },
            { title: "Duration", value: durationLabel(d.Execution_Duration) },
        ],
        highlight: { text: `🧵  ${str(d.Process_Name)} finished in ${durationLabel(d.Execution_Duration)}`, color: "accent" },
        actions: [
            {
                type: "Action.Submit",
                title: "🎫 Create Ticket",
                data: {
                    actionType: "createTicket",
                    alertType: "PROCESS",
                    Process_Name: d.Process_Name,
                    Process_Id: d.Process_Id,
                    Atom_Name: d.Atom_Name,
                    Execution_Type: d.Execution_Type,
                    Execution_Duration: d.Execution_Duration,
                },
            },
        ],
        sentBy,
    });
}
