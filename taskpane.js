// taskpane.js

// Authentication Handler Class
class AuthenticationHandler {
  constructor() {
    this.msalInstance = null;
    this.isInitialized = false;
    this.scopes = ["Mail.ReadWrite", "Mail.Send"];
    
    // Your specific tenant ID - replace with your actual tenant ID
    this.tenantId = "a05d37d6-5a0d-4083-887d-6b340796809a"; // e.g., "contoso.onmicrosoft.com" or a GUID
  }

  async initialize() {
    // Make sure MSAL is available
    if (typeof msal === 'undefined') {
      console.error("MSAL library not loaded");
      return false;
    }

    try {
      const msalConfig = {
        auth: {
          clientId: "f2ec0036-695b-419b-bbc7-fa83e14a7ccc", // Your client ID
          authority: `https://login.microsoftonline.com/${this.tenantId}`, // Specific tenant
          redirectUri: "https://smallcharbel.github.io/taskpane.html" // MUST EXACTLY match the URI in Azure portal
        },
        cache: {
          cacheLocation: "sessionStorage",
          storeAuthStateInCookie: true
        }
      };

      // Initialize MSAL instance
      this.msalInstance = new msal.PublicClientApplication(msalConfig);
      this.isInitialized = true;
      console.log("Authentication initialized successfully");
      return true;
    } catch (error) {
      console.error("Failed to initialize authentication:", error);
      return false;
    }
  }

  async getAccessToken() {
    if (!this.isInitialized || !this.msalInstance) {
      const initialized = await this.initialize();
      if (!initialized) {
        throw new Error("Authentication could not be initialized");
      }
    }

    try {
      const accounts = this.msalInstance.getAllAccounts();
      
      if (accounts.length > 0) {
        // Account exists, try silent token acquisition
        try {
          const silentRequest = {
            account: accounts[0],
            scopes: this.scopes
          };
          
          const response = await this.msalInstance.acquireTokenSilent(silentRequest);
          return response.accessToken;
        } catch (error) {
          // Silent acquisition failed, fall back to interactive method
          if (error instanceof msal.InteractionRequiredAuthError) {
            const interactiveRequest = {
              scopes: this.scopes
            };
            const response = await this.msalInstance.acquireTokenPopup(interactiveRequest);
            return response.accessToken;
          } else {
            throw error;
          }
        }
      } else {
        // No accounts, start interactive login
        const loginRequest = {
          scopes: this.scopes
        };
        await this.msalInstance.loginPopup(loginRequest);
        return this.getAccessToken(); // Try again now that we're logged in
      }
    } catch (error) {
      console.error("Authentication error:", error);
      throw error;
    }
  }
}

// Create a single instance of the auth handler
const authHandler = new AuthenticationHandler();

// Update status function
function updateStatus(message, type) {
  const statusContainer = document.getElementById("status-container");
  statusContainer.innerHTML = message;
  statusContainer.className = type || "";
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
    
    // Initialize authentication
    try {
      await authHandler.initialize();
    } catch (error) {
      console.error("Could not initialize authentication:", error);
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
    const accessToken = await authHandler.getAccessToken();
    
    // Get the current item
    const item = Office.context.mailbox.item;
    const messageId = item.itemId;
    
    // Call your Azure Function
    const functionUrl = "https://outlookaddintestptai.azurewebsites.net/api/forward-email?code=qZtLOtMh1tNugQdgNA20-2KnY0-2vIc9hkpamqw1c99bAzFudm7pyQ==";
    
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
      updateStatus("Email forwarded successfully with all attachments! (Test Version)", "success");
    } else {
      updateStatus("Error: " + (result.error || "Unknown error"), "error");
    }
  } catch (error) {
    updateStatus(`Error: ${error.message}`, "error");
    console.error("Error forwarding email:", error);
  }
}
