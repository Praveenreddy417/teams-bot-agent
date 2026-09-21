import axios from "axios";
import * as dotenv from "dotenv";
dotenv.config({ path: "./env/.env.local" });

const OTP_GEN_URL =
    process.env.OTP_GEN_URL ??
    "https://apibaseqa.easystepin.com/ws/rest/teams_bot/teams_login/otp_generation";
const OTP_VAL_URL =
    process.env.OTP_VAL_URL ??
    "https://apibaseqa.easystepin.com/ws/rest/teams_bot/OTP_validation/login_validation";

export interface OtpResult {
    success: boolean;
    message: string;
}

/** Safely read a field that may appear under several casing/spacing variants */
function field(obj: any, ...keys: string[]): string | undefined {
    if (!obj) return undefined;
    for (const k of keys) {
        if (obj[k] !== undefined && obj[k] !== null) return String(obj[k]);
    }
    return undefined;
}

function parseResponse(data: any): OtpResult {
    const code = field(data, "Status_code", "Status code", "status_code", "status code");
    const resp = field(data, "Status_Response", "Status Response", "status_response");
    const msg = field(
        data,
        "Status_Message", "Status message", "Status_message", "Status Message",
        "status_message", "status message"
    );
    const success = code === "200" || resp?.toLowerCase() === "success";
    return { success, message: msg ?? (success ? "Success." : "Something went wrong.") };
}

async function post(url: string, payload: Record<string, string>): Promise<OtpResult> {
    try {
        const resp = await axios.post(url, payload, {
            headers: { "Content-Type": "application/json" },
            timeout: 20_000,
        });
        const data = typeof resp.data === "string" ? JSON.parse(resp.data) : resp.data;
        return parseResponse(data);
    } catch (err: any) {
        const raw = err.response?.data;
        if (raw) {
            const data = typeof raw === "string" ? JSON.parse(raw) : raw;
            if (data) return parseResponse(data);
        }
        return { success: false, message: err.message ?? "Request failed." };
    }
}

export const requestOtp = (email: string) => post(OTP_GEN_URL, { Email_ID: email });
export const validateOtp = (email: string, otp: string) => post(OTP_VAL_URL, { Email_id: email, Otp: otp });