import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type {
  Channel,
  ChannelSummary,
  ChatMessage,
  ConversationMember,
  GraphApiResponse,
  MemberSummary,
  MessageSummary,
  Team,
  TeamSummary,
} from "../types/graph.js";

export function registerTeamsTools(server: McpServer, graphService: GraphService) {
  // List user's teams
  server.tool("list_teams", {}, async () => {
    try {
      const client = await graphService.getClient();
      const response = (await client.api("/me/joinedTeams").get()) as GraphApiResponse<Team>;

      if (!response?.value?.length) {
        return {
          content: [
            {
              type: "text",
              text: "No teams found.",
            },
          ],
        };
      }

      const teamList: TeamSummary[] = response.value.map((team: Team) => ({
        id: team.id,
        displayName: team.displayName,
        description: team.description,
        isArchived: team.isArchived,
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(teamList, null, 2),
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
  });

  // List channels in a team
  server.tool(
    "list_channels",
    {
      teamId: z.string().describe("Team ID"),
    },
    async ({ teamId }) => {
      try {
        const client = await graphService.getClient();
        const response = (await client
          .api(`/teams/${teamId}/channels`)
          .get()) as GraphApiResponse<Channel>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No channels found in this team.",
              },
            ],
          };
        }

        const channelList: ChannelSummary[] = response.value.map((channel: Channel) => ({
          id: channel.id,
          displayName: channel.displayName,
          description: channel.description,
          membershipType: channel.membershipType,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(channelList, null, 2),
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

  // Get channel messages
  server.tool(
    "get_channel_messages",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(20)
        .describe("Number of messages to retrieve (default: 20)"),
    },
    async ({ teamId, channelId, limit }) => {
      try {
        const client = await graphService.getClient();

        // Build query parameters - Teams channel messages API has limited query support
        // Only $top is supported, no $orderby, $filter, etc.
        const queryParams: string[] = [`$top=${limit}`];
        const queryString = queryParams.join("&");

        const response = (await client
          .api(`/teams/${teamId}/channels/${channelId}/messages?${queryString}`)
          .get()) as GraphApiResponse<ChatMessage>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No messages found in this channel.",
              },
            ],
          };
        }

        const messageList: MessageSummary[] = response.value.map((message: ChatMessage) => ({
          id: message.id,
          content: message.body?.content,
          from: message.from?.user?.displayName,
          createdDateTime: message.createdDateTime,
          importance: message.importance,
        }));

        // Sort messages by creation date (newest first) since API doesn't support orderby
        messageList.sort((a, b) => {
          const dateA = new Date(a.createdDateTime || 0).getTime();
          const dateB = new Date(b.createdDateTime || 0).getTime();
          return dateB - dateA;
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  totalReturned: messageList.length,
                  hasMore: !!response["@odata.nextLink"],
                  messages: messageList,
                },
                null,
                2
              ),
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

  // Send message to channel
  server.tool(
    "send_channel_message",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      message: z.string().describe("Message content"),
      importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
    },
    async ({ teamId, channelId, message, importance = "normal" }) => {
      try {
        const client = await graphService.getClient();

        const newMessage = {
          body: {
            content: message,
            contentType: "text",
          },
          importance: importance,
        };

        const result = (await client
          .api(`/teams/${teamId}/channels/${channelId}/messages`)
          .post(newMessage)) as ChatMessage;

        return {
          content: [
            {
              type: "text",
              text: `✅ Message sent successfully. Message ID: ${result?.id}`,
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

  // Get replies to a message in a channel
  server.tool(
    "get_channel_message_replies",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      messageId: z.string().describe("Message ID to get replies for"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(20)
        .describe("Number of replies to retrieve (default: 20)"),
    },
    async ({ teamId, channelId, messageId, limit }) => {
      try {
        const client = await graphService.getClient();

        // Only $top is supported for message replies
        const queryParams: string[] = [`$top=${limit}`];
        const queryString = queryParams.join("&");

        const response = (await client
          .api(
            `/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies?${queryString}`
          )
          .get()) as GraphApiResponse<ChatMessage>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No replies found for this message.",
              },
            ],
          };
        }

        const repliesList: MessageSummary[] = response.value.map((reply: ChatMessage) => ({
          id: reply.id,
          content: reply.body?.content,
          from: reply.from?.user?.displayName,
          createdDateTime: reply.createdDateTime,
          importance: reply.importance,
        }));

        // Sort replies by creation date (oldest first for replies)
        repliesList.sort((a, b) => {
          const dateA = new Date(a.createdDateTime || 0).getTime();
          const dateB = new Date(b.createdDateTime || 0).getTime();
          return dateA - dateB;
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  parentMessageId: messageId,
                  totalReplies: repliesList.length,
                  hasMore: !!response["@odata.nextLink"],
                  replies: repliesList,
                },
                null,
                2
              ),
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

  // Reply to a message in a channel
  server.tool(
    "reply_to_channel_message",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      messageId: z.string().describe("Message ID to reply to"),
      message: z.string().describe("Reply content"),
      importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
    },
    async ({ teamId, channelId, messageId, message, importance = "normal" }) => {
      try {
        const client = await graphService.getClient();

        const newReply = {
          body: {
            content: message,
            contentType: "text",
          },
          importance: importance,
        };

        const result = (await client
          .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`)
          .post(newReply)) as ChatMessage;

        return {
          content: [
            {
              type: "text",
              text: `✅ Reply sent successfully. Reply ID: ${result?.id}`,
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

  // List team members
  server.tool(
    "list_team_members",
    {
      teamId: z.string().describe("Team ID"),
    },
    async ({ teamId }) => {
      try {
        const client = await graphService.getClient();
        const response = (await client
          .api(`/teams/${teamId}/members`)
          .get()) as GraphApiResponse<ConversationMember>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No members found in this team.",
              },
            ],
          };
        }

        const memberList: MemberSummary[] = response.value.map((member: ConversationMember) => ({
          id: member.id,
          displayName: member.displayName,
          roles: member.roles,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(memberList, null, 2),
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
