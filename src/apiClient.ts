import axios, { AxiosInstance } from "axios";
import * as dotenv from "dotenv";

// Load environment variables path
dotenv.config({ path: "./env/.env.local" });

export interface HttpClientOptions {
    baseURL: string;
    authorization?: string;
    timeoutMs?: number;
    headers?: Record<string, string>;
}

/**
 * Creates a fully configured Axios instance with request/response interceptors.
 * Use this factory for every external API — never create raw axios instances elsewhere.
 */
export function createHttpClient(options: HttpClientOptions): AxiosInstance {
    const { baseURL, authorization, timeoutMs = 20000, headers = {} } = options;

    const client = axios.create({
        baseURL,
        headers: {
            "Content-Type": "application/json",
            ...(authorization ? { Authorization: authorization } : {}),
            ...headers,
        },
        timeout: timeoutMs,
        responseType: "text",
    });

    // ── Request interceptor ─────────────────────────────────────────────────
    client.interceptors.request.use(
        (config) => {
            (config as any).metadata = { startTime: Date.now() };

            const fullUrl = `${config.baseURL ?? ""}${config.url ?? ""}`;
            console.log("\n🚀 ═══════════════ API REQUEST ═══════════════");
            console.log("➡️  URL    :", fullUrl);
            console.log("➡️  Method :", config.method?.toUpperCase());
            console.log("➡️  Payload:", JSON.stringify(config.data, null, 2));
            console.log("═════════════════════════════════════════════\n");

            return config;
        },
        (error) => {
            console.error("REQUEST INTERCEPTOR ERROR:", error);
            return Promise.reject(error);
        }
    );

    // ── Response interceptor ────────────────────────────────────────────────
    client.interceptors.response.use(
        (response) => {
            const duration = Date.now() - ((response.config as any).metadata?.startTime ?? 0);
            const fullUrl = `${response.config.baseURL ?? ""}${response.config.url ?? ""}`;

            console.log("\n ═══════════════ API RESPONSE ══════════════");
            console.log("⬅️  URL      :", fullUrl);
            console.log("⬅️  Status   :", response.status);
            console.log("⏱️  Duration :", duration + "ms");
            console.log(
                "⬅️  Preview  :",
                typeof response.data === "string"
                    ? response.data.slice(0, 300)
                    : JSON.stringify(response.data, null, 2).slice(0, 300)
            );
            console.log("═════════════════════════════════════════════\n");

            return response;
        },
        (error) => {
            console.error("\n ═══════════════ API ERROR ═════════════════");
            console.error("➡️  URL    :", error.config?.url);
            console.error("➡️  Method :", error.config?.method?.toUpperCase());
            console.error("➡️  Status :", error.response?.status);
            console.error("➡️  Message:", error.message);
            console.error("═════════════════════════════════════════════\n");

            return Promise.reject(error);
        }
    );

    return client;
}

// ── Chat API Tunnel Client Instance ──────────────────────────────
export const pythonApi = createHttpClient({
    baseURL: process.env.BOT_PYTHON_API || "",
    timeoutMs: 0 // no timeout — chat backend can take longer than 30s
});

// ── Easy CICD QA Automation Client Instance ──────────────────────
export const boomiApi = createHttpClient({
    baseURL: (process.env.BOT_BOOMI_API || ""), // Added explicit trailing slash
    timeoutMs: 15000,
    authorization: process.env.API_BASIC_AUTH ? `Basic ${process.env.API_BASIC_AUTH}` : undefined
});