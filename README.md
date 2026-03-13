# Customer 360 — Dynamics 365

A single-page app that authenticates via MSAL and displays a Customer 360 view for any account in your Dynamics 365 Sales environment, showing **open cases**, **open opportunities**, and **contacts** — all from live Dataverse data.

---

## Prerequisites

| Requirement | Details |
|---|---|
| **Azure AD App Registration** | Register at [portal.azure.com](https://portal.azure.com) → Azure Active Directory → App registrations |
| **Redirect URI** | Add `http://localhost:3000` as a **Single-page application (SPA)** redirect URI |
| **API Permission** | Add **Dynamics CRM → `user_impersonation`** (Delegated) and grant admin consent |
| **Dynamics 365 environment** | `https://hyagom.crm.dynamics.com` (already configured) |

## Setup

1. **Register the Azure AD App** (if not already done):
   - Go to Azure Portal → App registrations → New registration.
   - Name: `Customer360` (or anything you like).
   - Supported account types: *Accounts in this organizational directory only*.
   - Redirect URI: select **SPA** and enter `http://localhost:3000`.

2. **Add API Permission**:
   - In your app registration → API permissions → Add a permission → Dynamics CRM → Delegated → `user_impersonation` → Add.
   - Click **Grant admin consent**.

3. **Update `config.js`**:
   - Set `clientId` to your Application (client) ID.
   - Set `tenantId` to your Directory (tenant) ID.

4. **Serve the app** (any static server on port 3000):
   ```bash
   # Using npx (no install required)
   npx http-server ./customer360 -p 3000

   # Or using Python
   cd customer360
   python -m http.server 3000
   ```

5. Open `http://localhost:3000` in a browser and sign in.

## File Structure

```
customer360/
├── index.html    — Main HTML page
├── styles.css    — Fluent-inspired styles
├── config.js     — MSAL + CRM configuration (edit this)
├── auth.js       — MSAL authentication logic
├── api.js        — Dataverse Web API calls
├── app.js        — UI rendering and event handling
└── README.md     — This file
```

## What It Shows

For a selected account:
- **KPI cards**: count of open cases, open opportunities, and contacts.
- **Open Cases table**: ticket number, title, priority, status, created date.
- **Opportunities table**: topic, estimated revenue, close date, sales stage, status.
- **Contacts table**: name, job title, email, phone.

All data comes directly from your Dynamics 365 environment via the Web API.
