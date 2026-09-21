import { pythonApi } from "./apiClient";
import { AdaptiveCardJson } from "./Chatcardbuilder";

export interface ChatResponse {
    session_id: string;
    type: "text" | "table" | "chart";
    mode?: "integration agent" | "integration custom" | string;
    answer?: string;
    data?: {
        title?: string;
        columns?: string[];
        rows?: string[][];
        chart_type?: string;
        card?: AdaptiveCardJson;
        // Present only when the requested chart type isn't supported in
        // Teams and the backend substituted the closest supported type.
        note?: string;
    };
}

const INTEGRATION_CHAT_URL =
    process.env.INTEGRATION_CHAT_URL ?? "https://aipass-flask.easystepin.com/esi_chat_bot/api/integrations/chat";

export async function askChat(
    question: string,
    sessionId: string | null,
    userEmail?: string
): Promise<ChatResponse> {
    try {
        console.log(`💬 Processing contextual query request`);
        console.log(`💬 User context: ${userEmail || "anonymous"}`);

        const resp = await pythonApi.post(INTEGRATION_CHAT_URL, {
            // email: userEmail || "",
            question,
            session_id: sessionId || ""
        });

        console.log("💬 askChat raw response:", resp.data);

        const data = typeof resp.data === "string" ? JSON.parse(resp.data) : resp.data;

        console.log("💬 askChat parsed response:", JSON.stringify(data));

        // ── Security & SQL Guard Response Mapper ─────────────────────────
        if (data?.error) {
            console.warn(`⚠️ Security policy triggered on AI core backend: ${data.error}`);
            return {
                session_id: sessionId || data.session_id || "",
                type: "text",
                answer: `⚠️ **Security Alert:** ${data.error}`,
            };
        }

        console.log(`💬 askChat parsed: type=${data?.type} mode=${data?.mode} keys=[${Object.keys(data ?? {}).join(", ")}]`);
        return data as ChatResponse;
    } catch (error: any) {
        console.error(`askChat routing failure:`, error.message);
        throw error;
    }
}