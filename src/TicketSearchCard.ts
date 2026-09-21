// Cards for the "/ticket" search & update flow.
import { TicketDetails, TicketUpdateResult, TicketCloseResult, TicketingTool } from "./TicketConnectorService";

export function buildTicketSearchCard(errorMessage?: string): object {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: "accent",
                bleed: true,
                items: [
                    {
                        type: "TextBlock",
                        text: "🎫 Search Ticket",
                        weight: "Bolder",
                        size: "Large",
                        spacing: "None",
                    },
                ],
            },
            ...(errorMessage
                ? [{ type: "TextBlock", text: errorMessage, color: "Attention", wrap: true, spacing: "Small" }]
                : []),
            { type: "TextBlock", text: "Ticketing Tool", weight: "Bolder", spacing: "Medium", wrap: true },
            {
                type: "Input.ChoiceSet",
                id: "tickting_tool",
                style: "compact",
                value: "ServiceNow",
                choices: [
                    { title: "ServiceNow", value: "ServiceNow" },
                    { title: "ADO", value: "ADO" },
                ],
            },
            { type: "TextBlock", text: "Incident Number", weight: "Bolder", spacing: "Medium", wrap: true },
            {
                type: "Input.Text",
                id: "incident_nb",
                placeholder: "INC0010064",
            },
        ],
        actions: [
            {
                type: "Action.Submit",
                title: "🔍 Search",
                style: "positive",
                data: { actionType: "searchTicket" },
            },
        ],
    };
}

export function buildTicketDetailCard(ticket: TicketDetails, tool: TicketingTool): object {
    const isOpen = ticket.status.toLowerCase() === "open";

    const facts = [
        { title: "Short Description", value: ticket.short_description || "N/A" },
        { title: "State", value: ticket.state || "N/A" },
        { title: "Priority", value: ticket.priority || "N/A" },
        { title: "Impact", value: ticket.impact || "N/A" },
        { title: "Urgency", value: ticket.urgency || "N/A" },
        { title: "Category", value: ticket.category || "N/A" },
        { title: "Opened At", value: ticket.opened_at || "N/A" },
        { title: "Status", value: ticket.status || "N/A" },
    ];

    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: isOpen ? "good" : "emphasis",
                bleed: true,
                items: [
                    {
                        type: "TextBlock",
                        text: `🎫 ${ticket.incident_number}`,
                        weight: "Bolder",
                        size: "Medium",
                        wrap: true,
                    },
                ],
            },
            {
                type: "TextBlock",
                text: ticket.description || "No description provided.",
                wrap: true,
                isSubtle: true,
                spacing: "Small",
            },
            { type: "FactSet", facts, spacing: "Medium" },
        ],
        actions: isOpen
            ? [
                  {
                      type: "Action.Submit",
                      title: "✏️ Update Ticket",
                      style: "positive",
                      data: { actionType: "openUpdateNotes", incident_nb: ticket.incident_number, tickting_tool: tool },
                  },
                  {
                      type: "Action.Submit",
                      title: "🔒 Close Ticket",
                      data: { actionType: "openCloseTicket", incident_nb: ticket.incident_number, tickting_tool: tool },
                  },
              ]
            : [],
    };
}

export function buildUpdateNotesCard(incident_nb: string, tool: TicketingTool): object {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: "accent",
                bleed: true,
                items: [
                    {
                        type: "TextBlock",
                        text: `✏️ Update ${incident_nb}`,
                        weight: "Bolder",
                        size: "Medium",
                        wrap: true,
                    },
                ],
            },
            { type: "TextBlock", text: "Work Notes", weight: "Bolder", spacing: "Medium", wrap: true },
            {
                type: "Input.Text",
                id: "work_notes",
                isMultiline: true,
                placeholder: "Currently analyzing the issue...",
            },
        ],
        actions: [
            {
                type: "Action.Submit",
                title: "✅ Submit Update",
                style: "positive",
                data: { actionType: "submitTicketUpdate", incident_nb, tickting_tool: tool },
            },
        ],
    };
}

export function buildCloseTicketCard(incident_nb: string, tool: TicketingTool): object {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: "attention",
                bleed: true,
                items: [
                    {
                        type: "TextBlock",
                        text: `🔒 Close ${incident_nb}`,
                        weight: "Bolder",
                        size: "Medium",
                        wrap: true,
                    },
                ],
            },
            { type: "TextBlock", text: "Close Code", weight: "Bolder", spacing: "Medium", wrap: true },
            {
                type: "Input.Text",
                id: "close_code",
                value: "Solution provided",
            },
            { type: "TextBlock", text: "Close Notes", weight: "Bolder", spacing: "Medium", wrap: true },
            {
                type: "Input.Text",
                id: "close_notes",
                isMultiline: true,
                placeholder: "Incident resolved successfully...",
            },
        ],
        actions: [
            {
                type: "Action.Submit",
                title: "🔒 Confirm Close",
                style: "destructive",
                data: { actionType: "submitTicketClose", incident_nb, tickting_tool: tool },
            },
        ],
    };
}

export function buildCloseResultCard(result: TicketCloseResult, incident_nb: string): object {
    const ok = result.success;
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
                        text: ok ? "🔒 Ticket Closed" : "⚠️ Close Failed",
                        weight: "Bolder",
                        size: "Medium",
                        wrap: true,
                    },
                ],
            },
            {
                type: "FactSet",
                spacing: "Medium",
                facts: [{ title: "Incident", value: result.incidentNumber ?? incident_nb }],
            },
            {
                type: "TextBlock",
                text: result.message,
                wrap: true,
                size: "Small",
                color: ok ? "good" : "attention",
                spacing: "Small",
            },
        ],
    };
}

export function buildUpdateResultCard(result: TicketUpdateResult, incident_nb: string): object {
    const ok = result.success;
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
                        text: ok ? "✅ Ticket Updated" : "⚠️ Update Failed",
                        weight: "Bolder",
                        size: "Medium",
                        wrap: true,
                    },
                ],
            },
            {
                type: "FactSet",
                spacing: "Medium",
                facts: [{ title: "Incident", value: result.incidentNumber ?? incident_nb }],
            },
            {
                type: "TextBlock",
                text: result.message,
                wrap: true,
                size: "Small",
                color: ok ? "good" : "attention",
                spacing: "Small",
            },
        ],
    };
}
