import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { DeviceCodeCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { initializeCachePlugin } from "../auth-plugin.js";
import { getAuthConfig } from "../config.js";

// Enable persistent token caching (stores refresh tokens)
initializeCachePlugin();

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
}

interface StoredAuthInfo {
  clientId: string;
  tenantId?: string;
  authenticated: boolean;
  timestamp: string;
  expiresAt?: string;
  token: string;
}

// Scopes for delegated (user) authentication
const DELEGATED_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
  "TeamMember.Read.All",
  "Chat.ReadBasic",
  "Chat.ReadWrite",
];

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private credential: DeviceCodeCredential | undefined;
  private readonly authPath = join(homedir(), ".msgraph-mcp-auth.json");
  private isInitialized = false;
  private authInfo: StoredAuthInfo | undefined;

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  private async initializeClient(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Check if we have stored auth info
      const authData = await fs.readFile(this.authPath, "utf8");
      this.authInfo = JSON.parse(authData);

      if (this.authInfo?.authenticated) {
        // Get config (use stored values if available, fall back to env vars)
        const config = getAuthConfig();
        const clientId = this.authInfo.clientId || config.clientId;
        const tenantId = this.authInfo.tenantId || config.tenantId;

        // Create credential with persistent cache - it will use cached refresh token
        this.credential = new DeviceCodeCredential({
          clientId,
          tenantId,
          tokenCachePersistenceOptions: {
            enabled: true,
            name: "teams-mcp-cache",
            unsafeAllowUnencryptedStorage: true,
          },
          // Silent callback since we're using cached tokens
          userPromptCallback: (info) => {
            console.error("Token refresh required. Please re-authenticate with: npx teams-mcp authenticate");
            console.error(`Visit: ${info.verificationUri}`);
            console.error(`Enter code: ${info.userCode}`);
          },
        });

        // Create Graph client using the credential
        this.client = Client.initWithMiddleware({
          authProvider: {
            getAccessToken: async () => {
              if (!this.credential) {
                throw new Error("No credential available");
              }
              // This will use cached refresh token to get new access token
              const token = await this.credential.getToken(DELEGATED_SCOPES);
              if (!token) {
                throw new Error("Failed to get access token");
              }
              return token.token;
            },
          },
        });

        this.isInitialized = true;
      }
    } catch (error) {
      // If no auth file exists, that's okay - just not authenticated
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
        console.error("Failed to initialize Graph client:", error);
      }
    }
  }

  async getAuthStatus(): Promise<AuthStatus> {
    await this.initializeClient();

    if (!this.client) {
      return { isAuthenticated: false };
    }

    try {
      const me = await this.client.api("/me").get();
      return {
        isAuthenticated: true,
        userPrincipalName: me?.userPrincipalName ?? undefined,
        displayName: me?.displayName ?? undefined,
        expiresAt: this.authInfo?.expiresAt,
      };
    } catch (error) {
      console.error("Error getting user info:", error);
      return { isAuthenticated: false };
    }
  }

  async getClient(): Promise<Client> {
    await this.initializeClient();

    if (!this.client) {
      throw new Error(
        "Not authenticated. Please run the authentication CLI tool first: npx teams-mcp authenticate"
      );
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }
}
