import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { Client } from "@microsoft/microsoft-graph-client";

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
}

interface StoredAuthInfo {
  clientId: string;
  authenticated: boolean;
  timestamp: string;
  expiresAt?: string;
  token: string;
}

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
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
      const authData = await fs.readFile(this.authPath, "utf8");
      this.authInfo = JSON.parse(authData);

      if (this.authInfo?.authenticated && this.authInfo?.token) {
        // Check if token is expired
        if (this.authInfo.expiresAt) {
          const expiresAt = new Date(this.authInfo.expiresAt);
          if (expiresAt <= new Date()) {
            console.log(
              "Token has expired. Please re-authenticate with: npx @floriscornel/teams-mcp@latest authenticate"
            );
            return;
          }
        }

        // Create Graph client with the saved token
        this.client = Client.initWithMiddleware({
          authProvider: {
            getAccessToken: async () => {
              if (!this.authInfo?.token) {
                throw new Error("No token available");
              }
              return this.authInfo.token;
            },
          },
        });

        this.isInitialized = true;
      }
    } catch (error) {
      console.error("Failed to initialize Graph client:", error);
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
        "Not authenticated. Please run the authentication CLI tool first: npx @floriscornel/teams-mcp@latest authenticate"
      );
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }
}
