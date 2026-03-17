/**
 * test-connection.js — Quick connectivity test for Dynamics 365
 * Authenticates via MSAL Device Code Flow and makes a test WhoAmI call.
 * Persists the token cache to disk so you only sign in once.
 */

const msal = require("@azure/msal-node");
const fetch = require("node-fetch");
const fs = require("fs");
const path = require("path");

const CRM_BASE_URL = "https://hyagom.crm.dynamics.com";
const API_URL = `${CRM_BASE_URL}/api/data/v9.2`;
const PUBLIC_CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const CACHE_FILE = path.join(__dirname, ".msal-cache.json");

function readCacheFromDisk() {
    if (fs.existsSync(CACHE_FILE)) {
        return fs.readFileSync(CACHE_FILE, "utf-8");
    }
    return null;
}

function writeCacheToDisk(cache) {
    fs.writeFileSync(CACHE_FILE, cache, "utf-8");
}

async function main() {
    const cacheData = readCacheFromDisk();

    const config = {
        auth: {
            clientId: PUBLIC_CLIENT_ID,
            authority: "https://login.microsoftonline.com/organizations",
        },
    };

    const pca = new msal.PublicClientApplication(config);

    // Restore cache from file if available
    if (cacheData) {
        pca.getTokenCache().deserialize(cacheData);
    }

    // Try to get cached accounts first
    const cache = pca.getTokenCache();
    const accounts = await cache.getAllAccounts();

    let tokenResponse;

    if (accounts.length > 0) {
        console.log(`Found cached account: ${accounts[0].username}`);
        try {
            tokenResponse = await pca.acquireTokenSilent({
                scopes: [`${CRM_BASE_URL}/.default`],
                account: accounts[0],
            });
            console.log("✔ Acquired token silently (no login needed).\n");
        } catch (e) {
            console.log("Silent token failed, falling back to device code...\n");
            tokenResponse = null;
        }
    }

    if (!tokenResponse) {
        tokenResponse = await pca.acquireTokenByDeviceCode({
            scopes: [`${CRM_BASE_URL}/.default`],
            deviceCodeCallback: (response) => {
                console.log("\n╔══════════════════════════════════════════════════════════╗");
                console.log("║  SIGN IN REQUIRED                                        ║");
                console.log("╚══════════════════════════════════════════════════════════╝");
                console.log(`\n${response.message}\n`);
            },
        });
        console.log(`✔ Authenticated as: ${tokenResponse.account.username}\n`);
    }

    // Persist cache to disk so future runs skip login
    writeCacheToDisk(pca.getTokenCache().serialize());

    // Test WhoAmI call
    console.log("→ Testing WhoAmI API call...");
    const res = await fetch(`${API_URL}/WhoAmI`, {
        headers: {
            Authorization: `Bearer ${tokenResponse.accessToken}`,
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
        },
    });

    if (!res.ok) {
        const err = await res.text();
        console.error(`✘ API call failed (${res.status}): ${err}`);
        process.exit(1);
    }

    const whoAmI = await res.json();
    console.log(`✔ Connected to Dynamics 365 successfully!`);
    console.log(`  Organization ID: ${whoAmI.OrganizationId}`);
    console.log(`  User ID:         ${whoAmI.UserId}`);
    console.log(`  Business Unit:   ${whoAmI.BusinessUnitId}`);
    console.log(`\n✔ Connection verified — you're all set!\n`);
}

main().catch((err) => {
    console.error("Error:", err.message);
    process.exit(1);
});
