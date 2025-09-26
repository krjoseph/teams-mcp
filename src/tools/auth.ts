import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";
import { ServerNotification, ServerRequest } from "@modelcontextprotocol/sdk/types.js";
import { RequestHandlerExtra } from "@modelcontextprotocol/sdk/shared/protocol.js";

export function registerAuthTools(server: McpServer, graphService: GraphService) {
  // Authentication status tool
  server.tool(
    "auth_status",
    "Check the authentication status of the Microsoft Graph connection. Returns whether the user is authenticated and shows their basic profile information.",
    {},
    { _meta: { requiredScopes: ["User.Read"] } } as any,
    async (_args: any, extra: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
      const status = await graphService.getAuthStatus(extra.requestInfo);
      return {
        content: [
          {
            type: "text",
            text: status.isAuthenticated
              ? `✅ Authenticated as ${status.displayName || "Unknown User"} (${status.userPrincipalName || "No email available"})`
              : "❌ Not authenticated. Please run: npx @floriscornel/teams-mcp@latest authenticate",
          },
        ],
      };
    }
  );
}
