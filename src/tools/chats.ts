import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type {
  Chat,
  ChatMessage,
  ChatSummary,
  ConversationMember,
  CreateChatPayload,
  GraphApiResponse,
  MessageSummary,
  User,
} from "../types/graph.js";

export function registerChatTools(server: McpServer, graphService: GraphService) {
  // List user's chats
  server.tool("list_chats", {}, async () => {
    try {
      const client = await graphService.getClient();
      const response = (await client.api("/me/chats").get()) as GraphApiResponse<Chat>;

      if (!response?.value?.length) {
        return {
          content: [
            {
              type: "text",
              text: "No chats found.",
            },
          ],
        };
      }

      const chatList: ChatSummary[] = response.value.map((chat: Chat) => ({
        id: chat.id,
        topic: chat.topic || "No topic",
        chatType: chat.chatType,
        memberCount: chat.members?.length,
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(chatList, null, 2),
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

  // Get chat messages
  server.tool(
    "get_chat_messages",
    {
      chatId: z.string().describe("Chat ID"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(20)
        .describe("Number of messages to retrieve"),
      since: z.string().optional().describe("Get messages since this ISO datetime"),
      until: z.string().optional().describe("Get messages until this ISO datetime"),
      fromUser: z.string().optional().describe("Filter messages from specific user ID"),
      orderBy: z
        .enum(["createdDateTime", "lastModifiedDateTime"])
        .optional()
        .default("createdDateTime")
        .describe("Sort order"),
      descending: z
        .boolean()
        .optional()
        .default(true)
        .describe("Sort in descending order (newest first)"),
    },
    async ({ chatId, limit, since, until, fromUser, orderBy, descending }) => {
      try {
        const client = await graphService.getClient();

        // Build query parameters
        const queryParams: string[] = [`$top=${limit}`];

        // Add ordering
        const sortDirection = descending ? "desc" : "asc";
        queryParams.push(`$orderby=${orderBy} ${sortDirection}`);

        // Add filters (only user filter is supported reliably)
        const filters: string[] = [];
        if (fromUser) {
          filters.push(`from/user/id eq '${fromUser}'`);
        }

        if (filters.length > 0) {
          queryParams.push(`$filter=${filters.join(" and ")}`);
        }

        const queryString = queryParams.join("&");

        const response = (await client
          .api(`/me/chats/${chatId}/messages?${queryString}`)
          .get()) as GraphApiResponse<ChatMessage>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No messages found in this chat with the specified filters.",
              },
            ],
          };
        }

        // Apply client-side date filtering since server-side filtering is not supported
        let filteredMessages = response.value;

        if (since || until) {
          filteredMessages = response.value.filter((message: ChatMessage) => {
            if (!message.createdDateTime) return true;

            const messageDate = new Date(message.createdDateTime);
            if (since) {
              const sinceDate = new Date(since);
              if (messageDate <= sinceDate) return false;
            }
            if (until) {
              const untilDate = new Date(until);
              if (messageDate >= untilDate) return false;
            }
            return true;
          });
        }

        const messageList: MessageSummary[] = filteredMessages.map((message: ChatMessage) => ({
          id: message.id,
          content: message.body?.content,
          from: message.from?.user?.displayName,
          createdDateTime: message.createdDateTime,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  filters: { since, until, fromUser },
                  filteringMethod: since || until ? "client-side" : "server-side",
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

  // Send chat message
  server.tool(
    "send_chat_message",
    {
      chatId: z.string().describe("Chat ID"),
      message: z.string().describe("Message content"),
      importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
      format: z
        .enum(["text", "markdown", "html"])
        .optional()
        .describe("Message format (text, markdown, html)"),
    },
    async ({ chatId, message, importance = "normal", format = "text" }) => {
      try {
        const client = await graphService.getClient();

        // Basic format validation and sanitization placeholder
        let contentType: "text" | "html" | "markdown" = "text";
        if (format === "html" || format === "markdown") {
          contentType = format;
          // TODO: Add sanitization/validation for HTML/Markdown
        }

        const newMessage = {
          body: {
            content: message,
            contentType: contentType,
          },
          importance: importance,
        };

        const result = (await client
          .api(`/me/chats/${chatId}/messages`)
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

  // Create new chat (1:1 or group)
  server.tool(
    "create_chat",
    {
      userEmails: z.array(z.string()).describe("Array of user email addresses to add to chat"),
      topic: z.string().optional().describe("Chat topic (for group chats)"),
    },
    async ({ userEmails, topic }) => {
      try {
        const client = await graphService.getClient();

        // Get current user ID
        const me = (await client.api("/me").get()) as User;

        // Create members array
        const members: ConversationMember[] = [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: {
              id: me?.id,
            },
            roles: ["owner"],
          } as ConversationMember,
        ];

        // Add other users as members
        for (const email of userEmails) {
          const user = (await client.api(`/users/${email}`).get()) as User;
          members.push({
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: {
              id: user?.id,
            },
            roles: ["member"],
          } as ConversationMember);
        }

        const chatData: CreateChatPayload = {
          chatType: userEmails.length === 1 ? "oneOnOne" : "group",
          members,
        };

        if (topic && userEmails.length > 1) {
          chatData.topic = topic;
        }

        const newChat = (await client.api("/chats").post(chatData)) as Chat;

        return {
          content: [
            {
              type: "text",
              text: `✅ Chat created successfully. Chat ID: ${newChat?.id}`,
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
