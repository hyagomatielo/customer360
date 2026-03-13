/**
 * inspect-schema.js — Connect to Dynamics 365 and fetch entity metadata
 * for Account, Case (incident), Opportunity, and Contact.
 */
const msal = require("@azure/msal-node");
const fetch = require("node-fetch");
const fs = require("fs");

const CRM_BASE_URL = "https://hyagom.crm.dynamics.com";
const API_URL = `${CRM_BASE_URL}/api/data/v9.2`;
const PUBLIC_CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

let accessToken = null;

async function authenticate() {
    const pca = new msal.PublicClientApplication({
        auth: { clientId: PUBLIC_CLIENT_ID, authority: "https://login.microsoftonline.com/organizations" },
    });
    const response = await pca.acquireTokenByDeviceCode({
        scopes: [`${CRM_BASE_URL}/.default`],
        deviceCodeCallback: (r) => console.log(`\n${r.message}\n`),
    });
    accessToken = response.accessToken;
    console.log(`Authenticated as: ${response.account.username}\n`);
}

async function crmGet(url) {
    const res = await fetch(`${API_URL}${url}`, {
        headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            Prefer: 'odata.include-annotations="*"',
        },
    });
    if (!res.ok) throw new Error(`${res.status}: ${await res.text()}`);
    return res.json();
}

async function getEntityAttributes(logicalName) {
    const data = await crmGet(
        `/EntityDefinitions(LogicalName='${logicalName}')/Attributes?$select=LogicalName,DisplayName,AttributeType,IsValidForRead,Description&$filter=IsValidForRead eq true&$orderby=LogicalName`
    );
    return data.value.map((a) => ({
        logicalName: a.LogicalName,
        displayName: a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName,
        type: a.AttributeType,
        description: a.Description?.UserLocalizedLabel?.Label || "",
    }));
}

async function getSampleRecords(entity, select, top = 3) {
    try {
        const data = await crmGet(`/${entity}?$select=${select}&$top=${top}`);
        return data.value;
    } catch (e) {
        return [];
    }
}

async function main() {
    await authenticate();

    const entities = ["account", "incident", "opportunity", "contact"];
    const result = {};

    for (const entity of entities) {
        console.log(`Fetching metadata for: ${entity}...`);
        result[entity] = await getEntityAttributes(entity);
        console.log(`  → ${result[entity].length} attributes found`);
    }

    // Also fetch sample data to understand what's populated
    console.log("\nFetching sample account data...");
    const sampleAccounts = await getSampleRecords(
        "accounts",
        "name,accountnumber,telephone1,telephone2,emailaddress1,websiteurl,address1_composite,address1_city,address1_stateorprovince,address1_country,revenue,numberofemployees,industrycode,description,ownerid,createdon,modifiedon,statuscode,statecode,primarycontactid",
        5
    );
    result.sampleAccounts = sampleAccounts;

    console.log("Fetching sample cases...");
    const sampleCases = await getSampleRecords(
        "incidents",
        "ticketnumber,title,prioritycode,statuscode,statecode,createdon,modifiedon,_customerid_value,casetypecode,severitycode,_ownerid_value",
        5
    );
    result.sampleCases = sampleCases;

    console.log("Fetching sample opportunities...");
    const sampleOpps = await getSampleRecords(
        "opportunities",
        "name,estimatedvalue,estimatedclosedate,salesstagecode,statuscode,statecode,closeprobability,_parentaccountid_value,stepname,_ownerid_value,actualvalue,budgetamount",
        5
    );
    result.sampleOpportunities = sampleOpps;

    console.log("Fetching sample contacts...");
    const sampleContacts = await getSampleRecords(
        "contacts",
        "fullname,jobtitle,emailaddress1,telephone1,mobilephone,_parentcustomerid_value,address1_city,department,createdon",
        5
    );
    result.sampleContacts = sampleContacts;

    // Get option set values for key fields
    console.log("Fetching option sets...");
    try {
        const industryMeta = await crmGet(
            "/EntityDefinitions(LogicalName='account')/Attributes(LogicalName='industrycode')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"
        );
        result.industryOptions = industryMeta.OptionSet?.Options?.map(o => ({
            value: o.Value,
            label: o.Label?.UserLocalizedLabel?.Label
        })) || [];
    } catch(e) { result.industryOptions = []; }

    try {
        const priorityMeta = await crmGet(
            "/EntityDefinitions(LogicalName='incident')/Attributes(LogicalName='prioritycode')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"
        );
        result.casePriorityOptions = priorityMeta.OptionSet?.Options?.map(o => ({
            value: o.Value,
            label: o.Label?.UserLocalizedLabel?.Label
        })) || [];
    } catch(e) { result.casePriorityOptions = []; }

    try {
        const statusMeta = await crmGet(
            "/EntityDefinitions(LogicalName='incident')/Attributes(LogicalName='statuscode')/Microsoft.Dynamics.CRM.StatusAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"
        );
        result.caseStatusOptions = statusMeta.OptionSet?.Options?.map(o => ({
            value: o.Value,
            label: o.Label?.UserLocalizedLabel?.Label
        })) || [];
    } catch(e) { result.caseStatusOptions = []; }

    try {
        const salesStageMeta = await crmGet(
            "/EntityDefinitions(LogicalName='opportunity')/Attributes(LogicalName='salesstagecode')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"
        );
        result.salesStageOptions = salesStageMeta.OptionSet?.Options?.map(o => ({
            value: o.Value,
            label: o.Label?.UserLocalizedLabel?.Label
        })) || [];
    } catch(e) { result.salesStageOptions = []; }

    const outputPath = require("path").join(__dirname, "schema-output.json");
    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
    console.log(`\nSchema saved to: ${outputPath}`);
}

main().catch((err) => {
    console.error("Failed:", err.message);
    process.exit(1);
});
