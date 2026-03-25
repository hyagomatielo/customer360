# Customer 360 for Dynamics 365

A standalone single-page application that authenticates via MSAL and provides two 360-degree views for Dynamics 365 Sales data.

## Pages

### Account Dashboard (`customer360.html`)
A full dashboard view for any Account record, featuring:
- **Account header** — name, phone, website, address, employees, annual revenue, status badge
- **KPI strip** — Open Cases, Opportunities, Pipeline Value, Contacts
- **Interactive D3.js charts** — Opportunity Pipeline (bar chart by stage), Cases by Priority (donut chart)
- **Data tables** — Opportunities, Cases, and Contacts with color-coded status badges
- Responsive layout adapting to different screen sizes

### Contact Profile (`contact360.html`)
A rich profile card for any Contact record, featuring:
- **Contact photo** with initials fallback and image upload
- **Loyalty tier** with animated shimmer effect (Diamond / Gold / Silver)
- **Contact details** — Customer ID, email, phone, address
- **Scores** — Lifetime Value, Propensity to Purchase, Engagement Score with progress bars
- **Relationship Health** — powered by Sales Insights (health status, trend, interactions)
- **Color customization** and inline editing

## Files

| File | Purpose |
|------|---------|
| `customer360.html` | Account dashboard page |
| `contact360.html` | Contact profile card page |
| `index.html` | Entry point / landing page |
| `styles.css` | Shared styles |

## Requirements

| Requirement | Details |
|---|---|
| **Azure AD App Registration** | Register at [portal.azure.com](https://portal.azure.com) → Azure Active Directory → App registrations |
| **Redirect URI** | Add `http://localhost:3000` as a **Single-page application (SPA)** redirect URI |
| **API Permission** | Add **Dynamics CRM → `user_impersonation`** (Delegated) and grant admin consent |
| **Dynamics 365 Sales** | Online environment with Sales Insights enabled for relationship health data |

## License

[MIT](LICENSE.txt)

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
