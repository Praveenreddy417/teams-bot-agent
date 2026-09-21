import * as fs from "fs";
import * as path from "path";

export interface AuthRecord {
    userId: string;
    email: string;
    authenticatedAt: string;
}

const STORE_FILE = path.join(__dirname, "..", "authStore.json");

function loadFromDisk(): Map<string, AuthRecord> {
    try {
        if (fs.existsSync(STORE_FILE)) {
            const raw = fs.readFileSync(STORE_FILE, "utf-8");
            const entries = JSON.parse(raw) as [string, AuthRecord][];
            return new Map(entries);
        }
    } catch (err) {
        console.warn("[AuthStore] Failed to load from disk, starting fresh:", err);
    }
    return new Map();
}

function saveToDisk(store: Map<string, AuthRecord>): void {
    try {
        fs.writeFileSync(
            STORE_FILE,
            JSON.stringify(Array.from(store.entries()), null, 2),
            "utf-8"
        );
    } catch (err) {
        console.error("[AuthStore] Failed to persist to disk:", err);
    }
}

class AuthStore {
    private store: Map<string, AuthRecord>;

    constructor() {
        this.store = loadFromDisk();
        console.log(`[AuthStore] Loaded ${this.store.size} authenticated user(s) from disk`);
    }

    login(userId: string, email: string): void {
        this.store.set(userId, { userId, email, authenticatedAt: new Date().toISOString() });
        saveToDisk(this.store);
        console.log(`[AuthStore] Logged in userId=${userId} email=${email}`);
    }

    logout(userId: string): void {
        this.store.delete(userId);
        saveToDisk(this.store);
        console.log(`[AuthStore] Logged out userId=${userId}`);
    }

    isAuthenticated(userId: string): boolean {
        return this.store.has(userId);
    }

    getEmail(userId: string): string | undefined {
        return this.store.get(userId)?.email;
    }
}

export const authStore = new AuthStore();