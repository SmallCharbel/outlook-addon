// taskpane.js

// Authentication configuration
const msalConfig = {
  auth: {
    clientId: "f2ec0036-695b-419b-bbc7-fa83e14a7ccc", // Your actual client ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + "/taskpane.html"
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: true
  }
};

// MSAL instance
const msalInstance = new msal.PublicClientApplication(msalConfig);

// Login Scope
const scopes = [
  "Mail.ReadWrite",
  "Mail.Send"
];

// Get Microsoft Graph token
async function getAccessToken() {
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

// Update status function
function updateStatus(message, type) {
  const statusContainer = document.getElementById("status-container");
  
  if (type === "processing") {
    statusContainer.innerHTML = `<span class="loading-icon"></span> ${message}`;
  } else if (type === "success") {
    statusContainer.innerHTML = `<img src="assets/success-icon.svg" width="16" height="16" alt="" style="margin-right: 8px;"> ${message}`;
  } else if (type === "error") {
    statusContainer.innerHTML = `<img src="assets/error-icon.svg" width="16" height="16" alt="" style="margin-right: 8px;"> ${message}`;
  } else {
    statusContainer.innerHTML = message;
  }
  
  statusContainer.className = `status-container ${type || ""}`;
  console.log(message);
}

// When Office is ready
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Add test version banner
    addTestVersionBanner();
    
    // Set up UI event handlers
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
  }
});

// Function to add test version banner
function addTestVersionBanner() {
  // Create test version banner
  const testBanner = document.createElement("div");
  testBanner.className = "test-version-banner";
  testBanner.innerHTML = "⚠️ TEST VERSION ⚠️";
  
  // Create test version description
  const testDescription = document.createElement("div");
  testDescription.className = "test-version-description";
  testDescription.innerHTML = 
    "This is a test version of the email forwarding add-in using Microsoft Graph API. " +
    "It includes attachment support and improved functionality.";
  
  // Find the content container and add the banner at the top
  const contentContainer = document.querySelector(".content-container") || document.body;
  contentContainer.insertBefore(testDescription, contentContainer.firstChild);
  contentContainer.insertBefore(testBanner, contentContainer.firstChild);
  
  // Add the CSS for the banner
  const style = document.createElement("style");
  style.textContent = `
    .test-version-banner {
      background-color: #FFC107;
      color: #000;
      text-align: center;
      padding: 8px;
      font-weight: bold;
      margin-bottom: 10px;
      border-radius: 4px;
    }
    
    .test-version-description {
      background-color: #f8f8f8;
      border-left: 4px solid #FFC107;
      padding: 10px;
      margin-bottom: 15px;
      font-size: 12px;
      color: #333;
    }
  `;
  document.head.appendChild(style);
}

// Function to forward email
async function forwardEmail() {
  updateStatus("Processing email...", "processing");
  
  try {
    // Get the access token - this will initialize auth if needed
    const accessToken = await getAccessToken();
    
    if (!accessToken) {
      updateStatus("Failed to obtain authentication token", "error");
      return;
    }
    
    console.log("Token obtained successfully", accessToken.substring(0, 10) + "...");
    
    // Get the current item
    const item = Office.context.mailbox.item;
    const messageId = item.itemId;
    
    if (!messageId) {
      updateStatus("Could not retrieve email ID", "error");
      return;
    }
    
    // Call your Azure Function
    const functionUrl = "https://outlookaddintestptai.azurewebsites.net/api/forward-email?code=qZtLOtMh1tNugQdgNA20-2KnY0-2vIc9hkpamqw1c99bAzFudm7pyQ==";
    
    updateStatus("Connecting to Azure Function...", "processing");
    
    console.log("Sending request to function with authorization header");
    
    const response = await fetch(functionUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${accessToken}`
      },
      body: JSON.stringify({
        messageId: messageId,
        accessToken: accessToken // Also send in body as fallback
      })
    });
    
    const responseText = await response.text();
    console.log("Response from function:", responseText);
    
    let result;
    try {
      result = JSON.parse(responseText);
    } catch (e) {
      console.error("Failed to parse response as JSON:", responseText);
      throw new Error("Invalid response from server");
    }
    
    if (!response.ok) {
      throw new Error(`Function returned status ${response.status}: ${responseText}`);
    }
    
    if (result.success) {
      updateStatus("Email forwarded successfully with all attachments! (Test Version)", "success");
    } else {
      updateStatus("Error: " + (result.error || "Unknown error"), "error");
    }
  } catch (error) {
    updateStatus(`Error: ${error.message}`, "error");
    console.error("Error forwarding email:", error);
  }
}
