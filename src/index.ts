#!/usr/bin/env node

import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import {
  PublicClientApplication,
  ConfidentialClientApplication,
  type AuthenticationResult,
  type Configuration,
} from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  getAuthConfig,
  logConfigInfo,
  validateConfig,
} from "./config.js";
import { cachePlugin } from "./msal-cache.js";
import { GraphService } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

const AUTH_INFO_PATH = join(homedir(), ".msgraph-mcp-auth.json");

// Authentication functions
async function authenticate() {
  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));

  try {
    let result: AuthenticationResult | null = null;

    if (config.isConfidentialClient && config.clientSecret) {
      // Confidential client flow (client credentials)
      console.log("\n🔒 Using client credentials flow...");

      const msalConfig: Configuration = {
        auth: {
          clientId: config.clientId,
          authority: config.authority,
          clientSecret: config.clientSecret,
        },
      };

      const client = new ConfidentialClientApplication(msalConfig);
      result = await client.acquireTokenByClientCredential({
        scopes: APP_SCOPE,
      });
    } else {
      // Public client flow (device code)
      console.log("\n📱 Using device code flow...");

      const msalConfig: Configuration = {
        auth: {
          clientId: config.clientId,
          authority: config.authority,
        },
        cache: {
          cachePlugin, // Use our custom file-based cache
        },
      };

      const client = new PublicClientApplication(msalConfig);

      result = await client.acquireTokenByDeviceCode({
        scopes: DELEGATED_SCOPES,
        deviceCodeCallback: (response) => {
          console.log("\n📱 Please complete authentication:");
          console.log(`🌐 Visit: ${response.verificationUri}`);
          console.log(`🔑 Enter code: ${response.userCode}`);
          console.log("\n⏳ Waiting for you to complete authentication...");
        },
      });
    }

    if (result) {
      // Save authentication info (for quick status checks)
      const authInfo = {
        clientId: CLIENT_ID,
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: result.expiresOn?.toISOString(),
        account: result.account?.username,
      };

      await fs.writeFile(AUTH_INFO_PATH, JSON.stringify(authInfo, null, 2));

      console.log("\n✅ Authentication successful!");
      console.log(`👤 Signed in as: ${result.account?.username || "Unknown"}`);
      console.log(`💾 Credentials saved to: ${AUTH_INFO_PATH}`);
      console.log("🔄 Refresh token cached for automatic renewal");
      console.log("\n🚀 You can now use the MCP server in Cursor!");
      console.log("   The server will automatically use these credentials.");
    }
  } catch (error) {
    console.error(
      "\n❌ Authentication failed:",
      error instanceof Error ? error.message : String(error)
    );
    process.exit(1);
  }
}

async function checkAuth() {
  try {
    const data = await fs.readFile(AUTH_INFO_PATH, "utf8");
    const authInfo = JSON.parse(data);

    if (authInfo.authenticated && authInfo.clientId) {
      console.log("✅ Authentication found");
      console.log(`👤 Account: ${authInfo.account || "Unknown"}`);
      console.log(`📅 Authenticated on: ${authInfo.timestamp}`);

      // Check if we have expiration info
      if (authInfo.expiresAt) {
        const expiresAt = new Date(authInfo.expiresAt);
        const now = new Date();

        if (expiresAt > now) {
          console.log(`⏰ Access token expires: ${expiresAt.toLocaleString()}`);
          console.log("🔄 Refresh token will automatically renew access");
          console.log("🎯 Ready to use with MCP server!");
        } else {
          console.log("⏰ Access token expired - will use refresh token");
          console.log("🎯 Ready to use with MCP server!");
        }
      } else {
        console.log("🎯 Ready to use with MCP server!");
      }
      return true;
    }
  } catch (_error) {
    console.log("❌ No authentication found");
    return false;
  }
  return false;
}

async function logout() {
  const CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.json");

  try {
    await fs.unlink(AUTH_INFO_PATH);
  } catch (_error) {
    // Ignore if file doesn't exist
  }

  try {
    await fs.unlink(CACHE_PATH);
  } catch (_error) {
    // Ignore if file doesn't exist
  }

  console.log("✅ Successfully logged out");
  console.log("🔄 Run 'npx @floriscornel/teams-mcp@latest authenticate' to re-authenticate");
}

// MCP Server setup
async function startMcpServer() {
  // Create MCP server
  const server = new McpServer({
    name: "teams-mcp",
    version: "0.3.3",
  });

  // Initialize Graph service (singleton)
  const graphService = GraphService.getInstance();

  // Register all tools
  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerTeamsTools(server, graphService);
  registerChatTools(server, graphService);
  registerSearchTools(server, graphService);

  // Start server
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Microsoft Graph MCP Server started");
}

// Main function to handle both CLI and MCP server modes
async function main() {
  const args = process.argv.slice(2);
  const command = args[0];

  // CLI commands
  switch (command) {
    case "authenticate":
    case "auth":
      await authenticate();
      return;
    case "check":
      await checkAuth();
      return;
    case "logout":
      await logout();
      return;
    case "help":
    case "--help":
    case "-h":
      console.log("Microsoft Graph MCP Server");
      console.log("");
      console.log("Usage:");
      console.log(
        "  npx @floriscornel/teams-mcp@latest authenticate # Authenticate with Microsoft"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest check        # Check authentication status"
      );
      console.log("  npx @floriscornel/teams-mcp@latest logout       # Clear authentication");
      console.log("  npx @floriscornel/teams-mcp@latest              # Start MCP server (default)");
      return;
    case undefined:
      // No command = start MCP server
      await startMcpServer();
      return;
    default:
      console.error(`Unknown command: ${command}`);
      console.error("Use --help to see available commands");
      process.exit(1);
  }
}

// Handle uncaught errors
process.on("uncaughtException", (error) => {
  console.error("Uncaught exception:", error);
  process.exit(1);
});

process.on("unhandledRejection", (reason, promise) => {
  console.error("Unhandled rejection at:", promise, "reason:", reason);
  process.exit(1);
});

main().catch((error) => {
  console.error("Failed to start:", error);
  process.exit(1);
});
