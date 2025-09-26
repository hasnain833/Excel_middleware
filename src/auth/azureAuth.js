const { ConfidentialClientApplication } = require("@azure/msal-node");
const logger = require("../config/logger");

class AzureAuthService {
  constructor() {
    this.clientApp = null;
    this.accessToken = null;
    this.tokenExpiry = null;
    this.initializeClient();
  }

  initializeClient() {
    try {
      const clientConfig = {
        auth: {
          clientId: process.env.AZURE_CLIENT_ID,
          clientSecret: process.env.AZURE_CLIENT_SECRET,
          authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        },
        system: {
          loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
              if (containsPii) return;
              logger.debug(`MSAL ${level}: ${message}`);
            },
            piiLoggingEnabled: false,
            logLevel: "Info",
          },
        },
      };

      this.clientApp = new ConfidentialClientApplication(clientConfig);
      logger.info("Azure AD client initialized successfully");
    } catch (error) {
      logger.error("Failed to initialize Azure AD client:", error);
      throw new Error("Azure AD initialization failed");
    }
  }

  async getAccessToken() {
    try {
      // Check if current token is still valid (with 5-minute buffer)
      if (
        this.accessToken &&
        this.tokenExpiry &&
        Date.now() < this.tokenExpiry - 5 * 60 * 1000
      ) {
        return this.accessToken;
      }

      const clientCredentialRequest = {
        scopes: ["https://graph.microsoft.com/.default"],
      };

      const response = await this.clientApp.acquireTokenByClientCredential(
        clientCredentialRequest
      );

      if (!response?.accessToken) {
        throw new Error("Failed to acquire access token");
      }

      // Cache the token and expiry time
      this.accessToken = response.accessToken;
      this.tokenExpiry = response.expiresOn.getTime();

      logger.info("Access token acquired successfully");
      return this.accessToken;
    } catch (error) {
      logger.error("Failed to acquire access token:", error.message);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }

  isTokenValid() {
    return (
      this.accessToken && this.tokenExpiry && Date.now() < this.tokenExpiry
    );
  }

  clearToken() {
    this.accessToken = null;
    this.tokenExpiry = null;
    logger.debug("Access token cache cleared");
  }

  getTokenInfo() {
    return {
      hasToken: !!this.accessToken,
      expiresAt: this.tokenExpiry
        ? new Date(this.tokenExpiry).toISOString()
        : null,
      isValid: this.isTokenValid(),
    };
  }
}

// Export singleton instance
module.exports = new AzureAuthService();
