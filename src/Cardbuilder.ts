import { BoomiProgress } from "./Boomiservice";
import { markdownToCardBody } from "./Markdowntocard";

export type AdaptiveCardJson = {
    $schema?: string;
    type: "AdaptiveCard";
    version: "1.5";
    body: any[];
    actions?: any[];
    [key: string]: any;
};

// ── Boomi Response Card ───────────────────────────────────────────────────────

export function buildBoomiResponseCard(
    content: string,
    progressSteps: BoomiProgress[],
    userQuestion: string
): AdaptiveCardJson {
    const actionSteps = progressSteps
        .filter((s) => s.type === "ACTION")
        .map((s) => s.content.replace(/^Calling Tool:\s*/i, "").trim());

    const generatedTitle = progressSteps.find((s) => s.type === "GENERATING_TITLE")?.content ?? "";

    // ── Header ───────────────────────────────────────────────────────
    const header: any = {
        type: "Container",
        style: "accent",
        bleed: true,
        items: [
            {
                type: "TextBlock",
                text: "📊 Integration Analytics",
                weight: "Bolder",
                size: "Large",
                spacing: "None",
            },
            ...(generatedTitle
                ? [{
                    type: "TextBlock",
                    text: generatedTitle,
                    size: "Small",
                    isSubtle: true,
                    spacing: "None",
                    wrap: true,
                }]
                : []),
        ],
    };

    // ── User question bubble ─────────────────────────────────────────
    const questionBlock: any = {
        type: "Container",
        style: "emphasis",
        spacing: "Small",
        items: [
            {
                type: "TextBlock",
                // text: `💬  ${userQuestion}`,
                wrap: true,
                isSubtle: true,
                size: "Small",
                spacing: "None",
            },
        ],
    };

    // ── Action steps (tools the agent called) ────────────────────────
    const actionRow: any[] = [];
    if (actionSteps.length > 0) {
        actionRow.push({
            type: "ColumnSet",
            spacing: "Small",
            columns: [
                {
                    type: "Column",
                    width: "auto",
                    items: [{
                        type: "TextBlock",
                        text: "⚙️",
                        size: "Small",
                        spacing: "None",
                    }],
                },
                {
                    type: "Column",
                    width: "stretch",
                    items: [{
                        type: "TextBlock",
                        text: actionSteps.join("  →  "),
                        size: "Small",
                        isSubtle: true,
                        wrap: true,
                        spacing: "None",
                    }],
                },
            ],
        });
    }

    // ── Content body ─────────────────────────────────────────────────
    const body = markdownToCardBody(content);

    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        msteams: { width: "Full" },
        body: [
            header,
            // questionBlock,
            ...actionRow,
            { type: "TextBlock", text: " ", separator: true, spacing: "Small" },
            ...body,
        ],
    };
}

// ── Greeting Card ─────────────────────────────────────────────────────────────

export function buildGreetingCard(name: string, email: string): AdaptiveCardJson {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        msteams: { width: "Full" },
        body: [
            {
                type: "Container",
                style: "accent",
                bleed: true,
                items: [
                    {
                        type: "TextBlock",
                        text: `👋 Hi, ${name}!`,
                        weight: "Bolder",
                        size: "Large",
                        spacing: "None",
                    },
                    {
                        type: "TextBlock",
                        text: "Integration Analytics",
                        size: "Small",
                        isSubtle: true,
                        spacing: "None",
                    },
                ],
            },
            // {
            //     type: "TextBlock",
            //     text: `Signed in as **${email}**`,
            //     size: "Small",
            //     isSubtle: true,
            //     spacing: "Small",
            // },
            {
                type: "TextBlock",
                text: "I'm your assistant for **integration data insights** — ask me about domain metrics, quarantine counts, golden records, model details, and more.",
                wrap: true,
                spacing: "Small",
            },
            // {
            //     type: "Container",
            //     style: "emphasis",
            //     spacing: "Medium",
            //     items: [
            //         {
            //             type: "TextBlock",
            //             text: "💡 **Try asking:**",
            //             weight: "Bolder",
            //             size: "Small",
            //             spacing: "None",
            //         },
            //         {
            //             type: "TextBlock",
            //             text: "• Get me all domain metrics\n• Show quarantine count for CONTACT domain\n• Give me 10 records of CONTRACT domain\n• What are the key insights for CUSTOMER domain?",
            //             wrap: true,
            //             size: "Small",
            //             spacing: "Small",
            //         },
            //     ],
            // },
            // {
            //     type: "TextBlock",
            //     text: "Type `/help` to see all commands.",
            //     size: "Small",
            //     isSubtle: true,
            //     spacing: "Small",
            // },
        ],
    };
}

// ── Help Card ─────────────────────────────────────────────────────────────────

export function buildHelpCard(email: string): AdaptiveCardJson {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        msteams: { width: "Full" },
        body: [
            {
                type: "Container",
                style: "accent",
                bleed: true,
                items: [
                    {
                        type: "TextBlock",
                        text: "📌 Integration Analytics — Help",
                        weight: "Bolder",
                        size: "Large",
                        spacing: "None",
                    },
                    // {
                    //     type: "TextBlock",
                    //     text: `Signed in as ${email}`,
                    //     size: "Small",
                    //     isSubtle: true,
                    //     spacing: "None",
                    // },
                ],
            },
            {
                type: "TextBlock",
                text: "Bot Commands",
                weight: "Bolder",
                spacing: "Medium",
            },
            {
                type: "FactSet",
                spacing: "Small",
                facts: [
                    { title: "hi", value: "Greet the bot" },
                    { title: "/help", value: "Show this list" },
                    { title: "/whoami", value: "Show your identity and session details" },
                    { title: "/ticket", value: "Search & update a ServiceNow ticket" },
                    { title: "/track", value: "Look up a tracked field value" },
                ],
            },
            {
                type: "TextBlock",
                text: "Sample Questions",
                weight: "Bolder",
                spacing: "Medium",
            },
            {
                type: "Container",
                style: "emphasis",
                spacing: "Small",
                items: [
                    {
                        type: "TextBlock",
                        text: "• Get me all domain metrics\n• Show quarantine count for all models\n• Give me 10 records of CONTACT domain\n• What is the golden record count for CONTRACT?\n• Show key insights for CUSTOMER domain",
                        wrap: true,
                        size: "Small",
                        spacing: "None",
                    },
                ],
            },
        ],
    };
}

// ── Who Am I Card ─────────────────────────────────────────────────────────────

export function buildWhoAmICard(
    name: string,
    email: string,
    userId: string,
    conversationType: string,
    conversationId: string
): AdaptiveCardJson {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "TextBlock",
                text: "👤 Your Identity",
                weight: "Bolder",
                size: "Large",
                wrap: true,
            },
            {
                type: "FactSet",
                spacing: "Medium",
                facts: [
                    { title: "Name", value: name || "N/A" },
                    { title: "Email", value: email || "Not available" },
                    { title: "User ID", value: userId },
                    { title: "Conversation Type", value: conversationType },
                    { title: "Conversation ID", value: conversationId },
                ],
            },
        ],
    };
}

// ── Error Card ────────────────────────────────────────────────────────────────

export function buildErrorCard(message: string): AdaptiveCardJson {
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
                        text: "⚠️ Something went wrong",
                        weight: "Bolder",
                        size: "Medium",
                        spacing: "None",
                    },
                ],
            },
            {
                type: "TextBlock",
                text: message,
                wrap: true,
                spacing: "Small",
            },
        ],
    };
}

// ── Login / OTP Cards ─────────────────────────────────────────────────────────

export function buildLoginCard(errorMessage?: string): AdaptiveCardJson {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: "attention",
                bleed: true,
                items: [{
                    type: "TextBlock",
                    text: "🔐 Login Required",
                    weight: "Bolder",
                    size: "Large",
                    spacing: "None",
                }],
            },
            {
                type: "TextBlock",
                text: "Please login to access **Integration Analytics**.",
                wrap: true,
                spacing: "Medium",
            },
            ...(errorMessage
                ? [{ type: "TextBlock", text: errorMessage, color: "Attention", wrap: true }]
                : []),
            { type: "TextBlock", text: "Email", weight: "Bolder", spacing: "Medium", wrap: true },
            { type: "Input.Text", id: "email", placeholder: "you@easystepin.com", style: "email" },
        ],
        actions: [
            {
                type: "Action.Submit",
                title: "📨 Send OTP",
                style: "positive",
                data: { actionType: "requestOtp" },
            },
        ],
    };
}

export function buildOtpCard(email: string, errorMessage?: string): AdaptiveCardJson {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: "accent",
                bleed: true,
                items: [{
                    type: "TextBlock",
                    text: "📩 Enter OTP",
                    weight: "Bolder",
                    size: "Large",
                    spacing: "None",
                }],
            },
            {
                type: "TextBlock",
                text: `OTP sent to **${email}**. Check your inbox.`,
                wrap: true,
                spacing: "Medium",
            },
            ...(errorMessage
                ? [{ type: "TextBlock", text: errorMessage, color: "Attention", wrap: true }]
                : []),
            { type: "TextBlock", text: "OTP", weight: "Bolder", spacing: "Medium", wrap: true },
            { type: "Input.Text", id: "otp", placeholder: "Enter OTP", maxLength: 6 },
        ],
        actions: [
            {
                type: "Action.Submit",
                title: "✅ Verify OTP",
                style: "positive",
                data: { actionType: "verifyOtp", email },
            },
            { type: "Action.Submit", title: "🔁 Resend OTP", data: { actionType: "resendOtp", email } },
            { type: "Action.Submit", title: "↩️ Use a different email", data: { actionType: "cancelLogin" } },
        ],
    };
}

export function buildLoginSuccessCard(email: string): AdaptiveCardJson {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: "good",
                bleed: true,
                items: [{
                    type: "TextBlock",
                    text: "✅ Logged In Successfully",
                    weight: "Bolder",
                    size: "Large",
                    spacing: "None",
                }],
            },
            {
                type: "TextBlock",
                text: `Welcome, **${email}**!\n\nYou can now ask me anything about your integration data. Type \`/help\` to get started.`,
                wrap: true,
                spacing: "Medium",
            },
        ],
    };
}

// ── Thinking card ─────────────────────────────────────────────────────────────

export function buildThinkingCard(step: string): AdaptiveCardJson {
    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [{ type: "TextBlock", text: `⏳ ${step}`, isSubtle: true, wrap: true }],
    };
}