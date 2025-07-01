import { describe, expect, it } from "vitest";
import type {
  ChannelSummary,
  ChatSummary,
  CreateChatPayload,
  GraphApiResponse,
  GraphError,
  MemberSummary,
  MessageFilterOptions,
  MessageSummary,
  RecentMessagesOptions,
  SearchHit,
  SearchRequest,
  SearchResponse,
  SendMessagePayload,
  TeamSummary,
  UserSummary,
} from "../graph.js";

describe("Graph Types", () => {
  describe("GraphApiResponse", () => {
    it("should define the correct structure", () => {
      const response: GraphApiResponse<{ id: string }> = {
        value: [{ id: "test" }],
        "@odata.count": 1,
        "@odata.nextLink": "https://example.com/next",
      };

      expect(response.value).toHaveLength(1);
      expect(response["@odata.count"]).toBe(1);
      expect(response["@odata.nextLink"]).toBe("https://example.com/next");
    });

    it("should allow optional properties", () => {
      const response: GraphApiResponse<{ id: string }> = {};

      expect(response.value).toBeUndefined();
      expect(response["@odata.count"]).toBeUndefined();
      expect(response["@odata.nextLink"]).toBeUndefined();
    });
  });

  describe("GraphError", () => {
    it("should define the correct structure", () => {
      const error: GraphError = {
        code: "InvalidRequest",
        message: "The request is invalid",
        innerError: {
          code: "BadRequest",
          message: "Bad request details",
          "request-id": "12345",
          date: "2023-01-01T00:00:00Z",
        },
      };

      expect(error.code).toBe("InvalidRequest");
      expect(error.message).toBe("The request is invalid");
      expect(error.innerError?.code).toBe("BadRequest");
    });
  });

  describe("UserSummary", () => {
    it("should allow all optional properties", () => {
      const user: UserSummary = {
        id: "user123",
        displayName: "John Doe",
        userPrincipalName: "john@example.com",
        mail: "john@example.com",
        jobTitle: "Developer",
        department: "Engineering",
        officeLocation: "Building A",
      };

      expect(user.id).toBe("user123");
      expect(user.displayName).toBe("John Doe");
    });

    it("should allow empty object", () => {
      const user: UserSummary = {};
      expect(user).toBeDefined();
    });
  });

  describe("TeamSummary", () => {
    it("should define team properties", () => {
      const team: TeamSummary = {
        id: "team123",
        displayName: "Engineering Team",
        description: "Software development team",
        isArchived: false,
      };

      expect(team.id).toBe("team123");
      expect(team.displayName).toBe("Engineering Team");
      expect(team.isArchived).toBe(false);
    });
  });

  describe("ChannelSummary", () => {
    it("should define channel properties", () => {
      const channel: ChannelSummary = {
        id: "channel123",
        displayName: "General",
        description: "General discussion",
        membershipType: "standard",
      };

      expect(channel.id).toBe("channel123");
      expect(channel.displayName).toBe("General");
    });
  });

  describe("ChatSummary", () => {
    it("should define chat properties", () => {
      const chat: ChatSummary = {
        id: "chat123",
        topic: "Project Discussion",
        chatType: "group",
        memberCount: 5,
      };

      expect(chat.id).toBe("chat123");
      expect(chat.topic).toBe("Project Discussion");
      expect(chat.chatType).toBe("group");
      expect(chat.memberCount).toBe(5);
    });
  });

  describe("MessageSummary", () => {
    it("should define message properties", () => {
      const message: MessageSummary = {
        id: "msg123",
        content: "Hello world",
        from: "John Doe",
        createdDateTime: "2023-01-01T10:00:00Z",
        importance: "normal",
      };

      expect(message.id).toBe("msg123");
      expect(message.content).toBe("Hello world");
      expect(message.from).toBe("John Doe");
    });
  });

  describe("MemberSummary", () => {
    it("should define member properties", () => {
      const member: MemberSummary = {
        id: "member123",
        displayName: "Jane Smith",
        roles: ["owner", "member"],
      };

      expect(member.id).toBe("member123");
      expect(member.displayName).toBe("Jane Smith");
      expect(member.roles).toEqual(["owner", "member"]);
    });
  });

  describe("CreateChatPayload", () => {
    it("should define chat creation structure", () => {
      const payload: CreateChatPayload = {
        chatType: "group",
        members: [
          {
            user: { id: "user1" },
            roles: ["owner"],
          } as any,
        ],
        topic: "New Project",
      };

      expect(payload.chatType).toBe("group");
      expect(payload.members).toHaveLength(1);
      expect(payload.topic).toBe("New Project");
    });

    it("should work without optional topic", () => {
      const payload: CreateChatPayload = {
        chatType: "oneOnOne",
        members: [],
      };

      expect(payload.chatType).toBe("oneOnOne");
      expect(payload.topic).toBeUndefined();
    });
  });

  describe("SendMessagePayload", () => {
    it("should define message sending structure", () => {
      const payload: SendMessagePayload = {
        body: {
          content: "Hello world",
          contentType: "text",
        },
        importance: "high",
      };

      expect(payload.body.content).toBe("Hello world");
      expect(payload.body.contentType).toBe("text");
      expect(payload.importance).toBe("high");
    });

    it("should work without optional importance", () => {
      const payload: SendMessagePayload = {
        body: {
          content: "Hello",
          contentType: "html",
        },
      };

      expect(payload.body.contentType).toBe("html");
      expect(payload.importance).toBeUndefined();
    });
  });

  describe("SearchRequest", () => {
    it("should define search request structure", () => {
      const request: SearchRequest = {
        entityTypes: ["chatMessage"],
        query: {
          queryString: "hello world",
        },
        from: 0,
        size: 25,
        enableTopResults: true,
      };

      expect(request.entityTypes).toEqual(["chatMessage"]);
      expect(request.query.queryString).toBe("hello world");
      expect(request.from).toBe(0);
      expect(request.size).toBe(25);
      expect(request.enableTopResults).toBe(true);
    });

    it("should work with minimal required properties", () => {
      const request: SearchRequest = {
        entityTypes: ["chatMessage"],
        query: {
          queryString: "test",
        },
      };

      expect(request.entityTypes).toEqual(["chatMessage"]);
      expect(request.query.queryString).toBe("test");
    });
  });

  describe("SearchResponse", () => {
    it("should define search response structure", () => {
      const response: SearchResponse = {
        value: [
          {
            searchTerms: ["hello"],
            hitsContainers: [
              {
                hits: [],
                total: 0,
                moreResultsAvailable: false,
              },
            ],
          },
        ],
      };

      expect(response.value).toHaveLength(1);
      expect(response.value[0].searchTerms).toEqual(["hello"]);
    });
  });

  describe("SearchHit", () => {
    it("should define search hit structure", () => {
      const hit: SearchHit = {
        hitId: "hit123",
        rank: 1,
        summary: "Found message",
        resource: {
          "@odata.type": "#microsoft.graph.chatMessage",
          id: "msg123",
          createdDateTime: "2023-01-01T10:00:00Z",
          from: {
            user: {
              displayName: "John Doe",
              id: "user123",
            },
          },
          body: {
            content: "Hello world",
            contentType: "text",
          },
          chatId: "chat123",
          channelIdentity: {
            teamId: "team123",
            channelId: "channel123",
          },
        },
      };

      expect(hit.hitId).toBe("hit123");
      expect(hit.rank).toBe(1);
      expect(hit.resource.id).toBe("msg123");
      expect(hit.resource.from?.user?.displayName).toBe("John Doe");
    });
  });

  describe("MessageFilterOptions", () => {
    it("should define filter options", () => {
      const options: MessageFilterOptions = {
        limit: 50,
        since: "2023-01-01T00:00:00Z",
        until: "2023-01-02T00:00:00Z",
        fromUser: "user123",
        mentionsUser: "user456",
        hasAttachments: true,
        importance: "high",
        search: "hello",
        orderBy: "createdDateTime",
      };

      expect(options.limit).toBe(50);
      expect(options.since).toBe("2023-01-01T00:00:00Z");
      expect(options.hasAttachments).toBe(true);
    });

    it("should work with empty options", () => {
      const options: MessageFilterOptions = {};
      expect(options).toBeDefined();
    });
  });

  describe("RecentMessagesOptions", () => {
    it("should extend MessageFilterOptions", () => {
      const options: RecentMessagesOptions = {
        limit: 25,
        includeChannels: true,
        includeChats: false,
        teamIds: ["team1", "team2"],
        chatIds: ["chat1", "chat2"],
        fromUser: "user123",
      };

      expect(options.limit).toBe(25);
      expect(options.includeChannels).toBe(true);
      expect(options.includeChats).toBe(false);
      expect(options.teamIds).toEqual(["team1", "team2"]);
      expect(options.chatIds).toEqual(["chat1", "chat2"]);
      expect(options.fromUser).toBe("user123");
    });
  });
});
