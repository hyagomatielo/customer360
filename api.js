/**
 * Dynamics 365 Web API Helper
 */

const API_BASE = `${APP_CONFIG.crmBaseUrl}/api/data/v9.2`;

/**
 * Generic fetch wrapper that adds the Bearer token and OData headers.
 */
async function crmFetch(relativeUrl) {
    const token = await getAccessToken();
    const response = await fetch(`${API_BASE}${relativeUrl}`, {
        headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            Prefer: 'odata.include-annotations="*"',
        },
    });

    if (!response.ok) {
        const errorBody = await response.text();
        throw new Error(`CRM API error ${response.status}: ${errorBody}`);
    }

    return response.json();
}

/**
 * Search accounts by name (top 10).
 */
async function searchAccounts(nameFilter) {
    const encoded = encodeURIComponent(nameFilter);
    const query = `/accounts?$select=name,accountnumber,telephone1,emailaddress1,address1_city&$filter=contains(name,'${encoded}')&$top=10&$orderby=name`;
    const data = await crmFetch(query);
    return data.value;
}

/**
 * Get open cases for an account.
 * statecode 0 = Active
 */
async function getOpenCases(accountId) {
    const query = `/incidents?$select=ticketnumber,title,prioritycode,statuscode,createdon&$filter=_customerid_value eq '${accountId}' and statecode eq 0&$orderby=createdon desc`;
    const data = await crmFetch(query);
    return data.value;
}

/**
 * Get open opportunities for an account.
 * statecode 0 = Open
 */
async function getOpenOpportunities(accountId) {
    const query = `/opportunities?$select=name,estimatedvalue,estimatedclosedate,salesstagecode,statuscode&$filter=_parentaccountid_value eq '${accountId}' and statecode eq 0&$orderby=estimatedclosedate asc`;
    const data = await crmFetch(query);
    return data.value;
}

/**
 * Get contacts for an account.
 */
async function getContacts(accountId) {
    const query = `/contacts?$select=fullname,jobtitle,emailaddress1,telephone1&$filter=_parentcustomerid_value eq '${accountId}'&$orderby=fullname`;
    const data = await crmFetch(query);
    return data.value;
}
