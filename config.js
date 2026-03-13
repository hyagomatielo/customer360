/**
 * ============================================================
 *  CONFIGURATION — UPDATE THESE VALUES BEFORE RUNNING
 * ============================================================
 *
 *  1. Register an app in Azure AD  (portal.azure.com → App registrations)
 *  2. Set a Redirect URI:  http://localhost:3000  (SPA type)
 *  3. Under API permissions, add:
 *        Dynamics CRM → user_impersonation  (Delegated)
 *  4. Copy the Application (client) ID and Tenant ID below.
 */

const APP_CONFIG = {
    // Replace with your Azure AD Application (client) ID
    clientId: "YOUR_CLIENT_ID_HERE",

    // Replace with your Azure AD Tenant ID (or "common" for multi-tenant)
    tenantId: "YOUR_TENANT_ID_HERE",

    // Your Dynamics 365 environment URL (no trailing slash)
    crmBaseUrl: "https://hyagom.crm.dynamics.com",

    // Redirect URI — must match the one registered in Azure AD
    redirectUri: "http://localhost:3000",

    // Scopes for Dynamics 365 Web API
    scopes: ["https://hyagom.crm.dynamics.com/.default"],
};
