import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type { GraphApiResponse, User, UserSummary } from "../types/graph.js";
import { RequestHandlerExtra } from "@modelcontextprotocol/sdk/shared/protocol.js";
import { ServerNotification, ServerRequest } from "@modelcontextprotocol/sdk/types.js";

export function registerUsersTools(server: McpServer, graphService: GraphService) {
  // Get current user
  server.tool(
    "get_current_user",
    "Get the current authenticated user's profile information including display name, email, job title, and department.",
    {},
    { _meta: { requiredScopes: ["User.Read"] } } as any,
    async (_args: any, extra: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
      try {
        const client = await graphService.getClient(extra.requestInfo);
        const user = (await client.api("/me").get()) as User;

        const userSummary: UserSummary = {
          displayName: user.displayName,
          userPrincipalName: user.userPrincipalName,
          mail: user.mail,
          id: user.id,
          jobTitle: user.jobTitle,
          department: user.department,
        };

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(userSummary, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Search users
  server.tool(
    "search_users",
    "Search for users in the organization by name or email address. Returns matching users with their basic profile information.",
    {
      query: z.string().describe("Search query (name or email)"),
    },
    { _meta: { requiredScopes: ["User.ReadBasic.All"] } } as any,
    async ({ query }, extra: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
      try {
        const client = await graphService.getClient(extra.requestInfo);
        const response = (await client
          .api("/users")
          .filter(
            `startswith(displayName,'${query}') or startswith(mail,'${query}') or startswith(userPrincipalName,'${query}')`
          )
          .get()) as GraphApiResponse<User>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No users found matching your search.",
              },
            ],
          };
        }

        const userList: UserSummary[] = response.value.map((user: User) => ({
          displayName: user.displayName,
          userPrincipalName: user.userPrincipalName,
          mail: user.mail,
          id: user.id,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(userList, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Get specific user
  server.tool(
    "get_user",
    "Get detailed information about a specific user by their ID or email address. Returns profile information including name, email, job title, and department.",
    {
      userId: z.string().describe("User ID or email address"),
    },
    { _meta: { requiredScopes: ["User.ReadBasic.All"] } } as any,
    async ({ userId }, extra: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
      try {
        const client = await graphService.getClient(extra.requestInfo);
        const user = (await client.api(`/users/${userId}`).get()) as User;

        const userSummary: UserSummary = {
          displayName: user.displayName,
          userPrincipalName: user.userPrincipalName,
          mail: user.mail,
          id: user.id,
          jobTitle: user.jobTitle,
          department: user.department,
          officeLocation: user.officeLocation,
        };

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(userSummary, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );
}
