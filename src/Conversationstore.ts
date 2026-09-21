import * as fs from "fs";
import * as path from "path";

export interface ConvRef {
    serviceUrl: string | undefined;
    conversationId: string;
    tenantId: string;
    botId: string;
    userId: string;
    userName: string;
    userEmail: string;
    conversationType: "personal" | "groupChat" | "channel";
    channelId?: string;
    teamId?: string;
}

const STORE_FILE = path.join(__dirname, "..", "conversationStore.json");

function loadFromDisk(): Map<string, ConvRef> {
    try {
        if (fs.existsSync(STORE_FILE)) {
            const raw = fs.readFileSync(STORE_FILE, "utf-8");
            const entries = JSON.parse(raw) as [string, ConvRef][];
            return new Map(entries);
        }
    } catch (err) {
        console.warn("[ConvStore] Failed to load from disk, starting fresh:", err);
    }
    return new Map();
}

function saveToDisk(store: Map<string, ConvRef>): void {
    try {
        fs.writeFileSync(
            STORE_FILE,
            JSON.stringify(Array.from(store.entries()), null, 2),
            "utf-8"
        );
    } catch (err) {
        console.error("[ConvStore] Failed to persist to disk:", err);
    }
}

class ConversationStore {
    private store: Map<string, ConvRef>;

    constructor() {
        this.store = loadFromDisk();
        console.log(`[ConvStore] Loaded ${this.store.size} conversation(s) from disk`);
    }

    save(userId: string, ref: ConvRef): void {
        this.store.set(userId, ref);
        saveToDisk(this.store);
    }

    saveGroup(conversationId: string, ref: ConvRef): void {
        this.store.set(conversationId, ref);
        saveToDisk(this.store);
    }

    getByEmail(email: string): ConvRef | undefined {
        const lower = email.toLowerCase();
        for (const ref of this.store.values()) {
            if (ref.conversationType === "personal" && ref.userEmail.toLowerCase() === lower)
                return ref;
        }
        return undefined;
    }

    getAllPersonal(): ConvRef[] {
        return Array.from(this.store.values()).filter((r) => r.conversationType === "personal");
    }

    getAllGroups(): ConvRef[] {
        return Array.from(this.store.values()).filter(
            (r) => r.conversationType === "groupChat" || r.conversationType === "channel"
        );
    }

    getAll(): ConvRef[] {
        return Array.from(this.store.values());
    }

    size(): number {
        return this.store.size;
    }

    delete(key: string): void {
        this.store.delete(key);
        saveToDisk(this.store);
    }
}

export const conversationStore = new ConversationStore();