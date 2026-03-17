/**
 * create-contact-fields.js — Create custom fields on the Contact entity
 * for the Customer 360 profile card.
 */
const msal = require("@azure/msal-node");
const fetch = require("node-fetch");
const fs = require("fs");
const path = require("path");

const CRM = "https://hyagom.crm.dynamics.com";
const API = CRM + "/api/data/v9.2";
const CACHE_FILE = path.join(__dirname, ".msal-cache.json");
const SOLUTION_NAME = "Customer360";

let accessToken = null;

async function authenticate() {
    const pca = new msal.PublicClientApplication({
        auth: { clientId: "51f81489-12ee-4a9e-aaae-a2591f45987d", authority: "https://login.microsoftonline.com/organizations" }
    });
    if (fs.existsSync(CACHE_FILE)) pca.getTokenCache().deserialize(fs.readFileSync(CACHE_FILE, "utf-8"));
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
        try {
            const tok = await pca.acquireTokenSilent({ scopes: [CRM + "/.default"], account: accounts[0] });
            accessToken = tok.accessToken;
            console.log("Authenticated silently as: " + accounts[0].username);
            return;
        } catch (e) { /* fall through */ }
    }
    const tok = await pca.acquireTokenByDeviceCode({
        scopes: [CRM + "/.default"],
        deviceCodeCallback: (r) => console.log(r.message),
    });
    accessToken = tok.accessToken;
    fs.writeFileSync(CACHE_FILE, pca.getTokenCache().serialize(), "utf-8");
}

async function crmRequest(method, url, body) {
    const res = await fetch(API + url, {
        method,
        headers: {
            Authorization: "Bearer " + accessToken,
            Accept: "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "MSCRM.SolutionUniqueName": SOLUTION_NAME,
        },
        body: body ? JSON.stringify(body) : undefined,
    });
    if (res.status === 204) return null;
    if (!res.ok) {
        const err = await res.text();
        return { error: err, status: res.status };
    }
    const text = await res.text();
    return text ? JSON.parse(text) : null;
}

// Get the contact entity metadata ID
async function getContactEntityId() {
    const res = await crmRequest("GET", "/EntityDefinitions(LogicalName='contact')?$select=MetadataId");
    return res.MetadataId;
}

// Create a String attribute
async function createStringAttribute(entityId, schemaName, displayName, maxLength, description) {
    console.log("  Creating: " + schemaName + " (" + displayName + ")...");
    const result = await crmRequest("POST", "/EntityDefinitions(" + entityId + ")/Attributes", {
        "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
        SchemaName: schemaName,
        DisplayName: { "@odata.type": "Microsoft.Dynamics.CRM.Label", LocalizedLabels: [{ "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: displayName, LanguageCode: 1033 }] },
        Description: { "@odata.type": "Microsoft.Dynamics.CRM.Label", LocalizedLabels: [{ "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: description, LanguageCode: 1033 }] },
        RequiredLevel: { Value: "None" },
        MaxLength: maxLength,
        FormatName: { Value: "Text" },
    });
    if (result && result.error) {
        if (result.error.includes("already exists")) { console.log("    (already exists)"); return; }
        console.error("    Error: " + result.error.substring(0, 200));
    } else {
        console.log("    OK");
    }
}

// Create a Money attribute
async function createMoneyAttribute(entityId, schemaName, displayName, description) {
    console.log("  Creating: " + schemaName + " (" + displayName + ")...");
    const result = await crmRequest("POST", "/EntityDefinitions(" + entityId + ")/Attributes", {
        "@odata.type": "Microsoft.Dynamics.CRM.MoneyAttributeMetadata",
        SchemaName: schemaName,
        DisplayName: { "@odata.type": "Microsoft.Dynamics.CRM.Label", LocalizedLabels: [{ "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: displayName, LanguageCode: 1033 }] },
        Description: { "@odata.type": "Microsoft.Dynamics.CRM.Label", LocalizedLabels: [{ "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: description, LanguageCode: 1033 }] },
        RequiredLevel: { Value: "None" },
        PrecisionSource: 2,
        ImeMode: "Disabled",
    });
    if (result && result.error) {
        if (result.error.includes("already exists")) { console.log("    (already exists)"); return; }
        console.error("    Error: " + result.error.substring(0, 200));
    } else {
        console.log("    OK");
    }
}

// Create an Integer attribute
async function createIntegerAttribute(entityId, schemaName, displayName, description, minValue, maxValue) {
    console.log("  Creating: " + schemaName + " (" + displayName + ")...");
    const result = await crmRequest("POST", "/EntityDefinitions(" + entityId + ")/Attributes", {
        "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
        SchemaName: schemaName,
        DisplayName: { "@odata.type": "Microsoft.Dynamics.CRM.Label", LocalizedLabels: [{ "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: displayName, LanguageCode: 1033 }] },
        Description: { "@odata.type": "Microsoft.Dynamics.CRM.Label", LocalizedLabels: [{ "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: description, LanguageCode: 1033 }] },
        RequiredLevel: { Value: "None" },
        MinValue: minValue,
        MaxValue: maxValue,
        Format: "None",
    });
    if (result && result.error) {
        if (result.error.includes("already exists")) { console.log("    (already exists)"); return; }
        console.error("    Error: " + result.error.substring(0, 200));
    } else {
        console.log("    OK");
    }
}

async function publishEntity() {
    console.log("\n  Publishing Contact entity...");
    const xml = '<importexportxml><entities><entity>contact</entity></entities></importexportxml>';
    const result = await crmRequest("POST", "/PublishXml", { ParameterXml: xml });
    if (result && result.error) {
        console.error("    Publish error: " + result.error.substring(0, 200));
    } else {
        console.log("    Published!");
    }
}

async function main() {
    await authenticate();

    console.log("\nFetching Contact entity metadata ID...");
    const entityId = await getContactEntityId();
    console.log("  Entity ID: " + entityId);

    console.log("\nCreating custom fields on Contact...");

    // Customer ID - string field
    await createStringAttribute(entityId, "cr4e0_CustomerID", "Customer ID", 50, "Unique customer identifier for profile card");

    // Loyalty Tier - string field
    await createStringAttribute(entityId, "cr4e0_LoyaltyTier", "Loyalty Tier", 50, "Customer loyalty tier (e.g., Gold, Silver, Bronze)");

    // Segment - string field
    await createStringAttribute(entityId, "cr4e0_Segment", "Segment", 100, "Customer segment classification");

    // Lifetime Value - money field
    await createMoneyAttribute(entityId, "cr4e0_LifetimeValue", "Lifetime Value", "Customer lifetime value");

    // Propensity to Purchase - integer (0-100)
    await createIntegerAttribute(entityId, "cr4e0_PropensityToPurchase", "Propensity to Purchase", "Propensity to purchase score (0-100%)", 0, 100);

    // Engagement Score - integer (0-100)
    await createIntegerAttribute(entityId, "cr4e0_EngagementScore", "Engagement Score", "Customer engagement score (0-100%)", 0, 100);

    await publishEntity();

    console.log("\nDone! Custom fields created on Contact entity.");
}

main().catch(e => { console.error("Error:", e.message); process.exit(1); });
