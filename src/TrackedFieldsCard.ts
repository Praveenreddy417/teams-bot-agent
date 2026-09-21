// Cards for the "/track" TrackedFields lookup flow.
import { TrackedFieldsResult } from "./TrackedFieldsService";

export function buildTrackedFieldsInputCard(errorMessage?: string): object {
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
                        text: "🔎 Track Field",
                        weight: "Bolder",
                        size: "Large",
                        spacing: "None",
                    },
                ],
            },
            ...(errorMessage
                ? [{ type: "TextBlock", text: errorMessage, color: "Attention", wrap: true, spacing: "Small" }]
                : []),
            { type: "TextBlock", text: "Process Name", weight: "Bolder", spacing: "Medium", wrap: true },
            {
                type: "Input.Text",
                id: "processname",
                placeholder: "Get_Users_Tracking_Process",
                isRequired: true,
                errorMessage: "Process name is required.",
            },
            { type: "TextBlock", text: "Tracking Field", weight: "Bolder", spacing: "Medium", wrap: true },
            {
                type: "Input.Text",
                id: "trackingField",
                placeholder: "e.g. an email, record ID, or status ID",
                isRequired: true,
                errorMessage: "Tracking field is required.",
            },
            { type: "TextBlock", text: "Connector", weight: "Bolder", spacing: "Medium", wrap: true },
            {
                type: "Input.ChoiceSet",
                id: "userselectconnector",
                style: "compact",
                value: "Source",
                isRequired: true,
                errorMessage: "Connector is required.",
                choices: [
                    { title: "Source", value: "Source" },
                    { title: "Target", value: "Target" },
                ],
            },
        ],
        actions: [
            {
                type: "Action.Submit",
                title: "🔍 Search",
                style: "positive",
                data: { actionType: "queryTrackedField" },
            },
        ],
    };
}

function formatValue(value: unknown): string {
    if (value === null || value === undefined || value === "") return "—";
    if (typeof value === "string") return value.replace(/\r?\n/g, " ").trim() || "—";
    return String(value);
}

export function buildTrackedFieldsResultCard(
    result: TrackedFieldsResult,
    payload: { processname: string; trackingField: string; userselectconnector: string }
): object {
    const ok = result.success && !!result.data;

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
                        text: ok ? "✅ Record Found" : "⚠️ No Record Found",
                        weight: "Bolder",
                        size: "Medium",
                        wrap: true,
                    },
                ],
            },
            {
                type: "FactSet",
                spacing: "Medium",
                facts: [
                    { title: "Process", value: payload.processname },
                    { title: "Tracking Field", value: payload.trackingField },
                    { title: "Connector", value: payload.userselectconnector },
                ],
            },
            ...(ok
                ? [
                      {
                          type: "FactSet",
                          spacing: "Medium",
                          facts: Object.entries(result.data!).map(([key, value]) => ({
                              title: key,
                              value: formatValue(value),
                          })),
                      },
                  ]
                : [
                      {
                          type: "TextBlock",
                          text: result.message,
                          wrap: true,
                          size: "Small",
                          color: "Attention",
                          spacing: "Small",
                      },
                  ]),
        ],
    };
}
