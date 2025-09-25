#!/usr/bin/env node

import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { DeviceCodeCredential } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { GraphService } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";
import { HttpTransportHandler } from "./transports/HttpTransportHandler.js";
import { StdioTransportHandler } from "./transports/StdioTransportHandler.js";

const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const TOKEN_PATH = join(homedir(), ".msgraph-mcp-auth.json");

// Types for command line arguments
interface ServerOptions {
  transport: 'stdio' | 'http';
  port?: number;
}

// Parse command line arguments
function parseArgs(args: string[]): { command: string | undefined; options: ServerOptions } {
  const options: ServerOptions = { transport: 'stdio' };
  let command: string | undefined;
  
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    
    if (arg === '--transport') {
      const transportValue = args[i + 1];
      if (transportValue === 'stdio' || transportValue === 'http') {
        options.transport = transportValue;
        i++; // Skip next argument as it's the value
      } else {
        console.error(`Invalid transport type: ${transportValue}. Must be 'stdio' or 'http'`);
        process.exit(1);
      }
    } else if (arg === '--port') {
      const portValue = parseInt(args[i + 1], 10);
      if (isNaN(portValue) || portValue < 1 || portValue > 65535) {
        console.error(`Invalid port number: ${args[i + 1]}. Must be between 1 and 65535`);
        process.exit(1);
      }
      options.port = portValue;
      i++; // Skip next argument as it's the value
    } else if (!arg.startsWith('--') && !command) {
      // First non-option argument is the command
      command = arg;
    }
  }
  
  return { command, options };
}

// Authentication functions
async function authenticate() {
  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));

  try {
    const credential = new DeviceCodeCredential({
      clientId: CLIENT_ID,
      tenantId: "common",
      userPromptCallback: (info) => {
        console.log("\n📱 Please complete authentication:");
        console.log(`🌐 Visit: ${info.verificationUri}`);
        console.log(`🔑 Enter code: ${info.userCode}`);
        console.log("\n⏳ Waiting for you to complete authentication...");
      },
    });

    // Get the actual token
    const token = await credential.getToken([
      "User.Read",
      "User.ReadBasic.All",
      "Team.ReadBasic.All",
      "Channel.ReadBasic.All",
      "ChannelMessage.Read.All",
      "ChannelMessage.Send",
      "TeamMember.Read.All",
      "Chat.ReadBasic",
      "Chat.ReadWrite",
    ]);

    if (token) {
      // Save authentication info with the actual token
      const authInfo = {
        clientId: CLIENT_ID,
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: token.expiresOnTimestamp
          ? new Date(token.expiresOnTimestamp).toISOString()
          : undefined,
        token: token.token, // Save the actual access token
      };

      await fs.writeFile(TOKEN_PATH, JSON.stringify(authInfo, null, 2));

      console.log("\n✅ Authentication successful!");
      console.log(`💾 Credentials saved to: ${TOKEN_PATH}`);
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
    const data = await fs.readFile(TOKEN_PATH, "utf8");
    const authInfo = JSON.parse(data);

    if (authInfo.authenticated && authInfo.clientId) {
      console.log("✅ Authentication found");
      console.log(`📅 Authenticated on: ${authInfo.timestamp}`);

      // Check if we have expiration info
      if (authInfo.expiresAt) {
        const expiresAt = new Date(authInfo.expiresAt);
        const now = new Date();

        if (expiresAt > now) {
          console.log(`⏰ Token expires: ${expiresAt.toLocaleString()}`);
          console.log("🎯 Ready to use with MCP server!");
        } else {
          console.log("⚠️  Token may have expired - please re-authenticate");
          return false;
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
  try {
    await fs.unlink(TOKEN_PATH);
    console.log("✅ Successfully logged out");
    console.log("🔄 Run 'npx @floriscornel/teams-mcp@latest authenticate' to re-authenticate");
  } catch (_error) {
    console.log("ℹ️  No authentication to clear");
  }
}

// MCP Server setup
async function startMcpServer(options: ServerOptions) {
  // Create MCP server
  const server = new McpServer(
    {
      name: "teams-mcp",
      version: "0.3.3",
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // Initialize Graph service (singleton)
  const graphService = GraphService.getInstance();

  // Register all tools
  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerTeamsTools(server, graphService);
  registerChatTools(server, graphService);
  registerSearchTools(server, graphService);

  // Start server with appropriate transport
  if (options.transport === 'http') {
    // Prioritize PORT environment variable, then command line argument, then default
    const port = process.env.PORT ? parseInt(process.env.PORT, 10) : options.port;
    const config = port ? { port } : {};
    const httpHandler = new HttpTransportHandler(server, config);
    await httpHandler.connect();
  } else {
    const stdioHandler = new StdioTransportHandler(server);
    await stdioHandler.connect();
  }
}

// Main function to handle both CLI and MCP server modes
async function main() {
  const args = process.argv.slice(2);
  
  // Check for help first before parsing
  if (args.includes('--help') || args.includes('-h') || args.includes('help')) {
    console.log("Microsoft Graph MCP Server");
    console.log("");
    console.log("Usage:");
    console.log(
      "  npx @floriscornel/teams-mcp@latest authenticate                    # Authenticate with Microsoft"
    );
    console.log(
      "  npx @floriscornel/teams-mcp@latest check                           # Check authentication status"
    );
    console.log(
      "  npx @floriscornel/teams-mcp@latest logout                          # Clear authentication"
    );
    console.log(
      "  npx @floriscornel/teams-mcp@latest [--transport stdio|http] [--port PORT] # Start MCP server (default)"
    );
    console.log("");
    console.log("Options:");
    console.log("  --transport stdio|http    Transport type (default: stdio)");
    console.log("  --port PORT              Port number for HTTP transport (default: 3000)");
    console.log("");
    console.log("Environment Variables:");
    console.log("  PORT                     Port number for HTTP transport (overrides --port)");
    return;
  }
  
  const { command, options } = parseArgs(args);

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
    case undefined:
      // No command = start MCP server
      await startMcpServer(options);
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
