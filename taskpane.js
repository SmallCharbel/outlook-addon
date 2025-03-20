// taskpane.js

let msalInstance = null; // Initialize variable at the top level

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sendMessageBtn").onclick = forwardEmail;
        
        // Display current user email
        document.getElementById("current-email").textContent = Office.context.mailbox.userProfile.emailAddress;
        
        // Get user initials for avatar
        const email = Office.context.mailbox.userProfile.emailAddress;
        if (email) {
            const nameParts = email.split('@')[0].split('.');
            let initials = "";
            if (nameParts.length >= 2) {
                initials = (nameParts[0].charAt(0) + nameParts[1].charAt(0)).toUpperCase();
            } else {
                initials = email.substring(0, 2).toUpperCase();
            }
            document.getElementById("user-initials").textContent = initials;
        }
        
        // Initialize MSAL after Office is ready
        initializeMsal();
    }
});

// Authentication configuration
function initializeMsal() {
    // Make sure the MSAL library is loaded
    if (typeof msal !== 'undefined') {
        const msalConfig = {
            auth: {
                clientId: "f2ec0036-695b-419b-bbc7-fa83e14a7ccc", // From your App Registration
                authority: "https://login.microsoftonline.com/common",
                redirectUri: window.location.origin + "/taskpane.html"
            },
            cache: {
                cacheLocation: "sessionStorage",
                storeAuthStateInCookie: true
            }
        };

        // Initialize MSAL instance
        msalInstance = new msal.PublicClientApplication(msalConfig);
        
        // Check if there's a cached account and silently acquire token
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            console.log("User already signed in");
        }
    } else {
        console.error("MSAL library not loaded. Make sure the script is included correctly.");
    }
}

function updateStatus(message, type) {
    const statusContainer = document.getElementById("status-container");
    statusContainer.innerHTML = message;
    statusContainer.className = type || "";
    console.log(message); // Log to console for debugging
}

// Login Scopes
const scopes = [
    "Mail.ReadWrite",
    "Mail.Send"
];

// Get Microsoft Graph token
async function getAccessToken() {
    if (!msalInstance) {
        throw new Error("Authentication not initialized. Please try again.");
    }
    
    try {
        const accounts = msalInstance.getAllAccounts();
        
        if (accounts.length > 0) {
            // Account exists, try silent token acquisition
            const silentRequest = {
                account: accounts[0],
                scopes: scopes
            };
            
            try {
                const response = await msalInstance.acquireTokenSilent(silentRequest);
                return response.accessToken;
            } catch (error) {
                // Silent acquisition failed, fall back to interactive method
                if (error instanceof msal.InteractionRequiredAuthError) {
                    const interactiveRequest = {
                        scopes: scopes
                    };
                    const response = await msalInstance.acquireTokenPopup(interactiveRequest);
                    return response.accessToken;
                } else {
                    throw error;
                }
            }
        } else {
            // No accounts, start interactive login
            const loginRequest = {
                scopes: scopes
            };
            const response = await msalInstance.loginPopup(loginRequest);
            return getAccessToken(); // Try again now that we're logged in
        }
    } catch (error) {
        console.error("Authentication error:", error);
        throw error;
    }
}

// Function to forward email
async function forwardEmail() {
    updateStatus("Processing email...", "processing");
    
    try {
        // Get the access token
        const accessToken = await getAccessToken();
        
        // Get the current item
        const item = Office.context.mailbox.item;
        const messageId = item.itemId;
        
        // Call your Azure Function - secure approach without exposing function key
        const functionUrl = "https://outlookaddintestptai.azurewebsites.net/api/forward-email?code=wwyxNq-WsRucsPjziT_7dD9l1NU5RJR_InSfZgsdFbwSAzFuCITcuA==";
        
        const response = await fetch(functionUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + accessToken
            },
            body: JSON.stringify({
                messageId: messageId
            })
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Function returned status ${response.status}: ${errorText}`);
        }
        
        const result = await response.json();
        
        if (result.success) {
            updateStatus("Email forwarded successfully with all attachments!", "success");
        } else {
            updateStatus("Error: " + (result.error || "Unknown error"), "error");
        }
    } catch (error) {
        updateStatus(`Error: ${error.message}`, "error");
        console.error("Error forwarding email:", error);
    }
}
