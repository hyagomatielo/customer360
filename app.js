/**
 * Application Logic — UI bindings and rendering
 */

let searchDebounceTimer = null;

/**
 * Called after successful MSAL authentication.
 */
function onAuthenticated(account) {
    document.getElementById("login-screen").classList.add("hidden");
    document.getElementById("main-content").classList.remove("hidden");
    document.getElementById("user-name").textContent = account.name || account.username;
}

/**
 * Debounced search handler.
 */
function onSearchInput() {
    clearTimeout(searchDebounceTimer);
    const query = document.getElementById("account-search").value.trim();
    const resultsDiv = document.getElementById("search-results");

    if (query.length < 2) {
        resultsDiv.classList.add("hidden");
        resultsDiv.innerHTML = "";
        return;
    }

    searchDebounceTimer = setTimeout(async () => {
        try {
            const accounts = await searchAccounts(query);
            renderSearchResults(accounts);
        } catch (err) {
            console.error("Search failed:", err);
            resultsDiv.innerHTML = '<div class="search-item error">Search failed. Check console.</div>';
            resultsDiv.classList.remove("hidden");
        }
    }, 350);
}

/**
 * Render the account search dropdown.
 */
function renderSearchResults(accounts) {
    const resultsDiv = document.getElementById("search-results");
    resultsDiv.innerHTML = "";

    if (accounts.length === 0) {
        resultsDiv.innerHTML = '<div class="search-item">No accounts found.</div>';
        resultsDiv.classList.remove("hidden");
        return;
    }

    accounts.forEach((acct) => {
        const item = document.createElement("div");
        item.className = "search-item";
        item.textContent = acct.name;
        if (acct.address1_city) {
            item.textContent += ` — ${acct.address1_city}`;
        }
        item.addEventListener("click", () => {
            selectAccount(acct);
            resultsDiv.classList.add("hidden");
            document.getElementById("account-search").value = acct.name;
        });
        resultsDiv.appendChild(item);
    });

    resultsDiv.classList.remove("hidden");
}

/**
 * Load all 360 data for the selected account.
 */
async function selectAccount(account) {
    const card = document.getElementById("account-card");
    card.classList.remove("hidden");

    // Header
    document.getElementById("account-name").textContent = account.name;
    document.getElementById("account-number").textContent = account.accountnumber
        ? `#${account.accountnumber}`
        : "";
    document.getElementById("account-phone").textContent = account.telephone1 || "";
    document.getElementById("account-email").textContent = account.emailaddress1 || "";
    document.getElementById("account-city").textContent = account.address1_city || "";

    // Reset KPIs
    document.getElementById("kpi-cases").textContent = "-";
    document.getElementById("kpi-opportunities").textContent = "-";
    document.getElementById("kpi-contacts").textContent = "-";

    // Show loaders
    showSection("cases", "loading");
    showSection("opps", "loading");
    showSection("contacts", "loading");

    const accountId = account.accountid;

    // Fetch all three in parallel
    const [cases, opps, contacts] = await Promise.all([
        getOpenCases(accountId).catch((e) => { console.error(e); return []; }),
        getOpenOpportunities(accountId).catch((e) => { console.error(e); return []; }),
        getContacts(accountId).catch((e) => { console.error(e); return []; }),
    ]);

    // Render Cases
    document.getElementById("kpi-cases").textContent = cases.length;
    renderCases(cases);

    // Render Opportunities
    document.getElementById("kpi-opportunities").textContent = opps.length;
    renderOpportunities(opps);

    // Render Contacts
    document.getElementById("kpi-contacts").textContent = contacts.length;
    renderContacts(contacts);
}

// ─── Section visibility helpers ─────────────────────────────

function showSection(prefix, state) {
    document.getElementById(`${prefix}-loading`).classList.toggle("hidden", state !== "loading");
    document.getElementById(`${prefix}-empty`).classList.toggle("hidden", state !== "empty");
    document.getElementById(`${prefix}-table`).classList.toggle("hidden", state !== "data");
}

// ─── Renderers ──────────────────────────────────────────────

const priorityLabels = { 1: "High", 2: "Normal", 3: "Low" };

function renderCases(cases) {
    if (cases.length === 0) { showSection("cases", "empty"); return; }
    showSection("cases", "data");

    const tbody = document.getElementById("cases-body");
    tbody.innerHTML = "";
    cases.forEach((c) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${escapeHtml(c.ticketnumber || "")}</td>
            <td>${escapeHtml(c.title || "")}</td>
            <td><span class="badge priority-${c.prioritycode}">${priorityLabels[c.prioritycode] || c.prioritycode}</span></td>
            <td>${escapeHtml(c["statuscode@OData.Community.Display.V1.FormattedValue"] || String(c.statuscode))}</td>
            <td>${formatDate(c.createdon)}</td>`;
        tbody.appendChild(tr);
    });
}

function renderOpportunities(opps) {
    if (opps.length === 0) { showSection("opps", "empty"); return; }
    showSection("opps", "data");

    const tbody = document.getElementById("opps-body");
    tbody.innerHTML = "";
    opps.forEach((o) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${escapeHtml(o.name || "")}</td>
            <td>${formatCurrency(o.estimatedvalue)}</td>
            <td>${formatDate(o.estimatedclosedate)}</td>
            <td>${escapeHtml(o["salesstagecode@OData.Community.Display.V1.FormattedValue"] || String(o.salesstagecode ?? ""))}</td>
            <td>${escapeHtml(o["statuscode@OData.Community.Display.V1.FormattedValue"] || String(o.statuscode))}</td>`;
        tbody.appendChild(tr);
    });
}

function renderContacts(contacts) {
    if (contacts.length === 0) { showSection("contacts", "empty"); return; }
    showSection("contacts", "data");

    const tbody = document.getElementById("contacts-body");
    tbody.innerHTML = "";
    contacts.forEach((ct) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${escapeHtml(ct.fullname || "")}</td>
            <td>${escapeHtml(ct.jobtitle || "")}</td>
            <td>${escapeHtml(ct.emailaddress1 || "")}</td>
            <td>${escapeHtml(ct.telephone1 || "")}</td>`;
        tbody.appendChild(tr);
    });
}

// ─── Utilities ──────────────────────────────────────────────

function escapeHtml(str) {
    const div = document.createElement("div");
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
}

function formatDate(dateStr) {
    if (!dateStr) return "";
    const d = new Date(dateStr);
    return d.toLocaleDateString();
}

function formatCurrency(value) {
    if (value == null) return "";
    return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value);
}

// Close dropdown when clicking outside
document.addEventListener("click", (e) => {
    const search = document.getElementById("account-search");
    const results = document.getElementById("search-results");
    if (!search.contains(e.target) && !results.contains(e.target)) {
        results.classList.add("hidden");
    }
});
