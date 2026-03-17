/**
 * deploy.js — Authenticate via MSAL Device Code Flow, then deploy
 * Customer 360 files as Web Resources into a Dynamics 365 solution.
 *
 * Usage:  node deploy.js
 */

const msal = require("@azure/msal-node");
const fetch = require("node-fetch");
const fs = require("fs");
const path = require("path");

// ─── Configuration ──────────────────────────────────────────────
const CRM_BASE_URL = "https://hyagom.crm.dynamics.com";
const API_URL = `${CRM_BASE_URL}/api/data/v9.2`;

// Well-known first-party Azure AD app ID for Dynamics 365
// This is the "Microsoft Power Platform CLI / PAC" public client
// that already has Dynamics CRM permissions. No app registration needed.
const PUBLIC_CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

const SOLUTION_UNIQUE_NAME = "Customer360";
const SOLUTION_DISPLAY_NAME = "Customer 360";
const SOLUTION_DESCRIPTION = "Customer 360 card showing open cases, opportunities, and contacts for an account.";
const PUBLISHER_PREFIX = "cr4e0"; // will be overridden by the default publisher's prefix

// Files to upload as web resources
const WEB_RESOURCE_FILES = [
    { file: "customer360.html", displayName: "Customer 360 - Dashboard", type: 1 }, // HTML (single-file embedded page)
    { file: "contact360.html", displayName: "Contact 360 - Profile Card", type: 1 }, // HTML (contact profile card)
];

const CACHE_FILE = path.join(__dirname, ".msal-cache.json");

let accessToken = null;

// ─── MSAL Authentication (Device Code Flow) ────────────────────

async function authenticate() {
    const config = {
        auth: {
            clientId: PUBLIC_CLIENT_ID,
            authority: "https://login.microsoftonline.com/organizations",
        },
    };

    const pca = new msal.PublicClientApplication(config);

    // Restore persistent cache
    if (fs.existsSync(CACHE_FILE)) {
        pca.getTokenCache().deserialize(fs.readFileSync(CACHE_FILE, "utf-8"));
    }

    // Try silent auth first
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
        try {
            const silentResponse = await pca.acquireTokenSilent({
                scopes: [`${CRM_BASE_URL}/.default`],
                account: accounts[0],
            });
            accessToken = silentResponse.accessToken;
            fs.writeFileSync(CACHE_FILE, pca.getTokenCache().serialize(), "utf-8");
            console.log(`✔ Authenticated silently as: ${accounts[0].username}\n`);
            return silentResponse;
        } catch (e) {
            console.log("  Silent auth failed, falling back to device code...\n");
        }
    }

    const deviceCodeRequest = {
        scopes: [`${CRM_BASE_URL}/.default`],
        deviceCodeCallback: (response) => {
            console.log("\n╔══════════════════════════════════════════════════════════╗");
            console.log("║  SIGN IN REQUIRED                                        ║");
            console.log("╚══════════════════════════════════════════════════════════╝");
            console.log(`\n${response.message}\n`);
        },
    };

    const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    accessToken = response.accessToken;
    fs.writeFileSync(CACHE_FILE, pca.getTokenCache().serialize(), "utf-8");
    console.log(`✔ Authenticated as: ${response.account.username}\n`);
    return response;
}

// ─── API Helpers ────────────────────────────────────────────────

async function crmRequest(method, relativeUrl, body) {
    const url = `${API_URL}${relativeUrl}`;
    const headers = {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Content-Type": "application/json",
    };

    const options = { method, headers };
    if (body) {
        options.body = JSON.stringify(body);
    }

    const res = await fetch(url, options);

    if (res.status === 204) return null;

    if (!res.ok) {
        const errText = await res.text();
        throw new Error(`API ${method} ${relativeUrl} → ${res.status}: ${errText}`);
    }

    // For POST/PATCH that return the entity ID in the header
    if (method === "POST" && res.headers.get("OData-EntityId")) {
        const entityIdUrl = res.headers.get("OData-EntityId");
        // Extract GUID from the URL
        const match = entityIdUrl.match(/\(([0-9a-f-]+)\)/i);
        return match ? match[1] : null;
    }

    const text = await res.text();
    return text ? JSON.parse(text) : null;
}

async function crmGet(url) {
    return crmRequest("GET", url);
}

async function crmPost(url, body) {
    return crmRequest("POST", url, body);
}

async function crmPatch(url, body) {
    return crmRequest("PATCH", url, body);
}

// ─── Get Default Publisher ──────────────────────────────────────

async function getDefaultPublisher() {
    // Try to find the "Default Publisher for <org>" (uniquename starts with "DefaultPublisherFor")
    const defaultPub = await crmGet("/publishers?$filter=startswith(uniquename,'DefaultPublisherFor')&$select=publisherid,uniquename,customizationprefix");
    if (defaultPub.value && defaultPub.value.length > 0) {
        return defaultPub.value[0];
    }

    // Fallback: get any non-system publisher (exclude MicrosoftCorporation and the system publisher)
    const allPubs = await crmGet("/publishers?$select=publisherid,uniquename,customizationprefix&$top=20&$orderby=createdon asc");
    if (allPubs.value) {
        const custom = allPubs.value.find(p =>
            p.uniquename !== "MicrosoftCorporation" &&
            p.uniquename !== "MicrosoftDynamics" &&
            !p.uniquename.startsWith("Microsft") &&
            !p.uniquename.startsWith("Microsoft") &&
            p.customizationprefix !== "none"
        );
        if (custom) return custom;
    }

    // Last resort: create our own publisher
    console.log("  → No custom publisher found. Creating one...");
    const newPubId = await crmPost("/publishers", {
        uniquename: "Customer360Publisher",
        friendlyname: "Customer 360 Publisher",
        customizationprefix: "c360",
        customizationoptionvalueprefix: 78600,
    });
    return { publisherid: newPubId, uniquename: "Customer360Publisher", customizationprefix: "c360" };
}

// ─── Solution Management ────────────────────────────────────────

async function findOrCreateSolution(publisherId) {
    console.log("→ Checking if solution already exists...");
    const existing = await crmGet(`/solutions?$filter=uniquename eq '${SOLUTION_UNIQUE_NAME}'&$select=solutionid,uniquename`);

    if (existing.value && existing.value.length > 0) {
        console.log(`  ✔ Solution "${SOLUTION_UNIQUE_NAME}" already exists.`);
        return existing.value[0].solutionid;
    }

    console.log(`→ Creating solution "${SOLUTION_DISPLAY_NAME}"...`);
    const solutionId = await crmPost("/solutions", {
        uniquename: SOLUTION_UNIQUE_NAME,
        friendlyname: SOLUTION_DISPLAY_NAME,
        description: SOLUTION_DESCRIPTION,
        version: "1.0.0.0",
        "publisherid@odata.bind": `/publishers(${publisherId})`,
    });

    console.log(`  ✔ Solution created (ID: ${solutionId})`);
    return solutionId;
}

// ─── Web Resource Management ────────────────────────────────────

function getWebResourceName(prefix, fileName) {
    // Convention: prefix_/customer360/filename
    return `${prefix}_/customer360/${fileName}`;
}

async function findWebResource(name) {
    const encoded = encodeURIComponent(name);
    const data = await crmGet(`/webresourceset?$filter=name eq '${encoded}'&$select=webresourceid,name`);
    if (data.value && data.value.length > 0) {
        return data.value[0];
    }
    return null;
}

async function createOrUpdateWebResource(prefix, fileDef) {
    const wrName = getWebResourceName(prefix, fileDef.file);
    const filePath = path.join(__dirname, fileDef.file);
    const content = fs.readFileSync(filePath, "utf-8");
    const contentBase64 = Buffer.from(content, "utf-8").toString("base64");

    const existing = await findWebResource(wrName);

    const wrBody = {
        name: wrName,
        displayname: fileDef.displayName,
        webresourcetype: fileDef.type,
        content: contentBase64,
    };

    if (existing) {
        console.log(`  ↻ Updating: ${wrName}`);
        await crmPatch(`/webresourceset(${existing.webresourceid})`, wrBody);
        return existing.webresourceid;
    } else {
        console.log(`  + Creating: ${wrName}`);
        const id = await crmPost("/webresourceset", wrBody);
        return id;
    }
}

// ─── Add Component to Solution ──────────────────────────────────

async function addWebResourceToSolution(webResourceId) {
    // Use the AddSolutionComponent action
    try {
        await crmPost("/AddSolutionComponent", {
            ComponentId: webResourceId,
            ComponentType: 61,  // Web Resource
            SolutionUniqueName: SOLUTION_UNIQUE_NAME,
            AddRequiredComponents: false,
            DoNotIncludeSubcomponents: false,
        });
    } catch (err) {
        // If already in solution, ignore the error
        if (err.message && err.message.includes("already exists")) {
            return;
        }
        // Some environments return an error if component is already there
        console.log(`  ⚠ Warning adding to solution: ${err.message.substring(0, 120)}`);
    }
}

// ─── Publish ────────────────────────────────────────────────────

async function publishAll() {
    console.log("\n→ Publishing all customizations...");
    await crmPost("/PublishAllXml", {});
    console.log("  ✔ Published successfully!\n");
}

// ─── Main ───────────────────────────────────────────────────────

async function main() {
    console.log("╔══════════════════════════════════════════════════════════╗");
    console.log("║  Customer 360 — Deploy to Dynamics 365                  ║");
    console.log("╚══════════════════════════════════════════════════════════╝\n");

    // Step 1: Authenticate
    console.log("Step 1/5: Authenticating...");
    await authenticate();

    // Step 2: Get publisher
    console.log("Step 2/5: Fetching publisher...");
    const publisher = await getDefaultPublisher();
    const prefix = publisher.customizationprefix;
    console.log(`  ✔ Using publisher "${publisher.uniquename}" (prefix: ${prefix})\n`);

    // Step 3: Create/find solution
    console.log("Step 3/5: Setting up solution...");
    await findOrCreateSolution(publisher.publisherid);
    console.log();

    // Step 4: Upload web resources
    console.log("Step 4/5: Uploading web resources...");
    const webResourceIds = [];
    for (const fileDef of WEB_RESOURCE_FILES) {
        const wrId = await createOrUpdateWebResource(prefix, fileDef);
        webResourceIds.push(wrId);
    }
    console.log();

    // Step 5: Add to solution
    console.log("Adding web resources to solution...");
    for (const wrId of webResourceIds) {
        await addWebResourceToSolution(wrId);
    }
    console.log("  ✔ All components added to solution.\n");

    // Step 6: Publish
    console.log("Step 5/5: Publishing...");
    await publishAll();

    console.log("══════════════════════════════════════════════════════════");
    console.log("  DEPLOYMENT COMPLETE!");
    console.log(`  Solution: ${SOLUTION_DISPLAY_NAME}`);
    console.log(`  Environment: ${CRM_BASE_URL}`);
    console.log(`  Open Dynamics 365 → Settings → Solutions to see it.`);
    console.log("══════════════════════════════════════════════════════════\n");
}

main().catch((err) => {
    console.error("\n✖ Deployment failed:", err.message);
    process.exit(1);
});
