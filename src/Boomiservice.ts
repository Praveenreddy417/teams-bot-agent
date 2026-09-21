import * as https from "https";
import * as http from "http";
import * as dotenv from "dotenv";
dotenv.config({ path: "./env/.env.local" });

const BOOMI_BASE_URL =
    process.env.BOOMI_BASE_URL ??
    "https://c01-usa-east.ai-agent-garden.boomi.com/execute/api/v1/easystepinitservicesllc-4ZNYGO";
const BOOMI_AUTH = process.env.BOOMI_AUTH ?? "";
const BOOMI_DEPLOYMENT_ID = process.env.BOOMI_DEPLOYMENT_ID ?? "237f6f95-f644-44a8-b21f-c1f38856cc0c";

export interface BoomiProgress {
    type: "THINKING" | "ACTION" | "GENERATING_TITLE" | string;
    content: string;
}

export interface BoomiResponse {
    sessionId: string;   // Retained conversation boundary reference token
    agentId: string;
    content: string;
    progressSteps: BoomiProgress[];
}

interface SseEvent {
    event: string;
    data: string;
}

/** * Parse raw multi-line stream text allocations without losing text across split network packets
 */
function parseSseChunk(raw: string, partial: string): { events: SseEvent[]; leftover: string } {
    const full = partial + raw;
    const blocks = full.split("\n\n");
    const leftover = blocks.pop() ?? "";

    const events: SseEvent[] = [];

    for (const block of blocks) {
        if (!block.trim()) continue;
        let event = "";
        let dataLines: string[] = [];

        for (const line of block.split("\n")) {
            if (line.startsWith("event:")) {
                event = line.substring(6).trim();
            } else if (line.startsWith("data:")) {
                dataLines.push(line.substring(5).trim());
            } else if (line.startsWith("id:") || line.startsWith(":")) {
                // Keep structural compliance filters active
            } else if (line.trim() && !event && dataLines.length === 0) {
                event = line.trim();
            }
        }

        const data = dataLines.join("\n");
        if (event || data) {
            events.push({ event: event || "data", data });
        }
    }

    return { events, leftover };
}

function safeJson(raw: string): any | null {
    try {
        if (!raw || raw === "null") return null;
        return JSON.parse(raw);
    } catch {
        return null;
    }
}

function buildPayload(question: string, sessionId: string | null): Record<string, string> {
    if (sessionId) {
        return {
            deployment_id: BOOMI_DEPLOYMENT_ID,
            session_id: sessionId,
            message: question,
        };
    }
    return {
        deployment_id: BOOMI_DEPLOYMENT_ID,
        message: question,
    };
}

/**
 * Execution hook linking the microservice architecture contexts to downstream engines
 */
export function askBoomiAgent(
    question: string,
    userEmail: string,
    sessionId: string | null = null,
    timeoutMs = 60_000
): Promise<BoomiResponse> {
    return new Promise((resolve, reject) => {
        const returnedSessionId: { value: string } = { value: "" };
        const agentId: { value: string } = { value: "" };
        const progressSteps: BoomiProgress[] = [];
        let finalContent = "";
        let partial = "";

        const payloadObj = buildPayload(question, sessionId);
        const payloadBuf = Buffer.from(JSON.stringify(payloadObj));

        console.log("\n🚀 ============== BOOMI API REQUEST ==============");
        console.log("➡️  URL      :", `${BOOMI_BASE_URL}/session`);
        console.log("➡️  Session  :", sessionId ? `CONTINUING (${sessionId})` : "NEW SESSION");
        console.log("=================================================\n");

        const url = new URL(`${BOOMI_BASE_URL}/session`);
        const protocol = url.protocol === "https:" ? https : http;

        const options: http.RequestOptions = {
            hostname: url.hostname,
            port: url.port || (url.protocol === "https:" ? 443 : 80),
            path: url.pathname,
            method: "POST",
            headers: {
                "Authorization": BOOMI_AUTH,
                "Content-Type": "application/json",
                "Accept": "text/event-stream",
                "Content-Length": payloadBuf.byteLength,
            },
        };

        const timer = setTimeout(() => {
            req.destroy();
            reject(new Error(`Boomi stream communication timed out past threshold (${timeoutMs}ms)`));
        }, timeoutMs);

        const req = protocol.request(options, (res) => {
            if (res.statusCode && res.statusCode >= 400) {
                clearTimeout(timer);
                let body = "";
                res.on("data", (c) => (body += c));
                res.on("end", () => reject(new Error(`Boomi HTTP Error State [${res.statusCode}]: ${body}`)));
                return;
            }

            res.setEncoding("utf-8");

            res.on("data", (chunk: string) => {
                const { events, leftover } = parseSseChunk(chunk, partial);
                partial = leftover;

                for (const { event, data } of events) {
                    if (event === "start") {
                        const parsed = safeJson(data);
                        if (parsed) {
                            returnedSessionId.value = parsed.session_id ?? "";
                            agentId.value = parsed.agent_id ?? "";
                        }
                        continue;
                    }

                    if (event === "progress_notification") {
                        const parsed = safeJson(data);
                        if (parsed?.type && parsed?.content) {
                            progressSteps.push({ type: parsed.type, content: parsed.content });
                        }
                        continue;
                    }

                    if (event === "message") {
                        const parsed = safeJson(data);
                        if (parsed?.role === "agent" && parsed?.content) {
                            // Check if incoming structure is a block rewrite or text fragment append target
                            if (finalContent && parsed.content.startsWith(finalContent)) {
                                finalContent = parsed.content;
                            } else {
                                finalContent += parsed.content;
                            }
                        }
                        continue;
                    }
                }
            });

            res.on("end", () => {
                clearTimeout(timer);

                const result: BoomiResponse = {
                    sessionId: returnedSessionId.value || sessionId || "",
                    agentId: agentId.value,
                    content: finalContent.trim() || "No analytical response returned from execution engines.",
                    progressSteps,
                };

                resolve(result);
            });

            res.on("error", (err) => {
                clearTimeout(timer);
                reject(err);
            });
        });

        req.on("error", (err) => {
            clearTimeout(timer);
            reject(err);
        });

        req.write(payloadBuf);
        req.end();
    });
}