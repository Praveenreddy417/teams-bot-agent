// Run this once (ts-node or compile) to find your bot's catalog app id.
// Usage: npx ts-node find-teams-app-id.ts "ESI Boomi"   (partial displayName match)
import * as dotenv from "dotenv";
dotenv.config({ path: "./env/.env.local" });

async function main() {
    const tenantId = process.env.AAD_TENANT_ID ?? process.env.TENANT_ID!;
    const clientId = process.env.AAD_CLIENT_ID ?? process.env.CLIENT_ID!;
    const clientSecret = process.env.AAD_CLIENT_SECRET ?? process.env.CLIENT_SECRET!;

    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            scope: "https://graph.microsoft.com/.default",
            grant_type: "client_credentials",
        }),
    });
    const { access_token } = await tokenRes.json();

    const res = await fetch("https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$expand=appDefinitions", {
        headers: { Authorization: `Bearer ${access_token}` },
    });
    const data = await res.json();

    if (!res.ok || !Array.isArray(data.value)) {
        console.error(`\n❌ Graph call failed (HTTP ${res.status}):`);
        console.error(JSON.stringify(data, null, 2));
        return;
    }

    const search = (process.argv[2] ?? "").toLowerCase();
    const matches = data.value.filter((a: any) =>
        !search || a.displayName?.toLowerCase().includes(search)
    );

    for (const app of matches) {
        console.log(`\n${app.displayName}`);
        console.log(`  teamsApp.id (this is TEAMS_APP_ID): ${app.id}`);
        console.log(`  distributionMethod: ${app.distributionMethod}`);
        for (const def of app.appDefinitions ?? []) {
            console.log(`  version ${def.version} — publishingState: ${def.publishingState}`);
        }
    }
    if (matches.length === 0) console.log("No matching apps found in catalog.");
}

main().catch(console.error);