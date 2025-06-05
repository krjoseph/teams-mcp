import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";

export function registerAuthTools(server: McpServer, graphService: GraphService) {
  // Authentication status tool
  server.tool("auth_status", {}, async () => {
    const status = await graphService.getAuthStatus();
    return {
      content: [
        {
          type: "text",
          text: status.isAuthenticated
            ? `✅ Authenticated as ${status.displayName} (${status.userPrincipalName})`
            : "❌ Not authenticated. Please run: npx @floriscornel/teams-mcp@latest authenticate",
        },
      ],
    };
  });
}
