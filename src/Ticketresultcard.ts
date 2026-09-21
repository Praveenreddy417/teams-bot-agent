// Confirmation card shown after the "Create Ticket" button is pressed.
import { TicketPayload, TicketResult } from "./Ticketservice";

export function buildTicketResultCard(
    result: TicketResult,
    payload: TicketPayload,
    createdBy?: string
): object {
    const ok = result.success;
    const timestamp = new Date().toLocaleString("en-IN", {
        timeZone: "Asia/Kolkata",
        dateStyle: "medium",
        timeStyle: "short",
    });

    const facts: { title: string; value: string }[] = [
        { title: "Priority", value: `P${payload.priority}` },
    ];
    if (result.ticketId) facts.unshift({ title: "Ticket ID", value: result.ticketId });
    if (createdBy) facts.push({ title: "Raised by", value: createdBy });

    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: ok ? "good" : "attention",
                bleed: true,
                items: [
                    {
                        type: "TextBlock",
                        text: ok ? "🎫  Ticket Created" : "⚠️  Ticket Creation Failed",
                        weight: "Bolder",
                        size: "Medium",
                        wrap: true,
                    },
                    {
                        type: "TextBlock",
                        text: `${timestamp} IST`,
                        size: "Small",
                        isSubtle: true,
                        spacing: "None",
                    },
                ],
            },
            {
                type: "TextBlock",
                text: payload.title,
                weight: "Bolder",
                wrap: true,
                spacing: "Medium",
            },
            {
                type: "TextBlock",
                text: payload.description,
                wrap: true,
                isSubtle: true,
                spacing: "Small",
            },
            { type: "FactSet", facts, spacing: "Medium" },
            ...(result.message
                ? [{
                    type: "TextBlock",
                    text: ok ? result.message : `Reason: ${result.message}`,
                    wrap: true,
                    size: "Small",
                    color: ok ? "good" : "attention",
                    spacing: "Small",
                }]
                : []),
        ],
    };
}
