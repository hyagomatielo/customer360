/**
 * MSAL Authentication Module
 */

const msalConfig = {
    auth: {
        clientId: APP_CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${APP_CONFIG.tenantId}`,
        redirectUri: APP_CONFIG.redirectUri,
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
    scopes: APP_CONFIG.scopes,
};

/**
 * Sign the user in via popup and return the account.
 */
async function signIn() {
    try {
        const response = await msalInstance.loginPopup(loginRequest);
        msalInstance.setActiveAccount(response.account);
        onAuthenticated(response.account);
    } catch (error) {
        console.error("Login failed:", error);
        alert("Sign-in failed. Check the console for details and verify your config.js values.");
    }
}

/**
 * Sign the user out.
 */
function signOut() {
    msalInstance.logoutPopup();
}

/**
 * Acquire an access token silently, falling back to popup if needed.
 */
async function getAccessToken() {
    const account = msalInstance.getActiveAccount();
    if (!account) {
        throw new Error("No active account. Please sign in first.");
    }

    const tokenRequest = {
        scopes: APP_CONFIG.scopes,
        account: account,
    };

    try {
        const response = await msalInstance.acquireTokenSilent(tokenRequest);
        return response.accessToken;
    } catch (error) {
        // Fallback to interactive if silent fails
        const response = await msalInstance.acquireTokenPopup(tokenRequest);
        return response.accessToken;
    }
}

// Check if user is already signed in on page load
(function handleRedirect() {
    msalInstance.handleRedirectPromise().then((response) => {
        if (response) {
            msalInstance.setActiveAccount(response.account);
            onAuthenticated(response.account);
        } else {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                msalInstance.setActiveAccount(accounts[0]);
                onAuthenticated(accounts[0]);
            }
        }
    });
})();
