# Teams MCP

A Model Context Protocol (MCP) server that provides seamless integration with Microsoft Graph APIs, enabling AI assistants to interact with Microsoft Teams, users, and organizational data.

## 🚀 Features

### 🔐 Authentication
- OAuth 2.0 authentication flow with Microsoft Graph
- Secure token management and refresh
- Authentication status checking

### 👥 User Management
- Get current user information
- Search users by name or email
- Retrieve detailed user profiles
- Access organizational directory data

### 🏢 Microsoft Teams Integration
- **Teams Management**
  - List user's joined teams
  - Access team details and metadata
  
- **Channel Operations**
  - List channels within teams
  - Retrieve channel messages
  - Send messages to team channels
  - Support for message importance levels (normal, high, urgent)
  
- **Team Members**
  - List team members and their roles
  - Access member information

### 💬 Chat & Messaging
- **1:1 and Group Chats**
  - List user's chats
  - Create new 1:1 or group conversations
  - Retrieve chat message history with filtering and pagination
  - Send messages to existing chats

### 🔍 Advanced Search & Discovery
- **Message Search**
  - Search across all Teams channels and chats using Microsoft Search API
  - Support for KQL (Keyword Query Language) syntax
  - Filter by sender, mentions, attachments, importance, and date ranges
  - Get recent messages with advanced filtering options
  - Find messages mentioning specific users



## 📦 Installation

```bash
# Install dependencies
bun install

# Build the project
bun run build

# Set up authentication
bun run auth
```

## 🔧 Configuration

### Prerequisites
- Node.js 18+ or Bun 1.0+
- Microsoft 365 account with appropriate permissions
- Azure App Registration with Microsoft Graph permissions

### Required Microsoft Graph Permissions
- `User.Read` - Read user profile
- `User.ReadBasic.All` - Read basic user info
- `Team.ReadBasic.All` - Read team information
- `Channel.ReadBasic.All` - Read channel information
- `ChannelMessage.Read.All` - Read channel messages
- `ChannelMessage.Send` - Send channel messages
- `Chat.Read` - Read chat messages
- `Chat.ReadWrite` - Create and manage chats
- `Mail.Read` - Required for Microsoft Search API
- `Calendars.Read` - Required for Microsoft Search API
- `Files.Read.All` - Required for Microsoft Search API
- `Sites.Read.All` - Required for Microsoft Search API

## 🛠️ Usage

### Starting the Server
```bash
# Development mode with hot reload
bun run dev

# Production mode
bun run build && node dist/index.js
```

### Available MCP Tools

#### Authentication
- `authenticate` - Initiate OAuth authentication flow
- `logout` - Clear authentication tokens
- `get_current_user` - Get authenticated user information

#### User Operations
- `search_users` - Search for users by name or email
- `get_user` - Get detailed user information by ID or email

#### Teams Operations
- `list_teams` - List user's joined teams
- `list_channels` - List channels in a specific team
- `get_channel_messages` - Retrieve messages from a team channel with pagination and filtering
- `send_channel_message` - Send a message to a team channel
- `list_team_members` - List members of a specific team

#### Chat Operations
- `list_chats` - List user's chats (1:1 and group)
- `get_chat_messages` - Retrieve messages from a specific chat with pagination and filtering
- `send_chat_message` - Send a message to a chat
- `create_chat` - Create a new 1:1 or group chat

#### Search Operations
- `search_messages` - Search across all Teams messages using KQL syntax
- `get_recent_messages` - Get recent messages with advanced filtering options
- `get_my_mentions` - Find messages mentioning the current user


## 📋 Examples

### Authentication

First, authenticate with Microsoft Graph:

```bash
npx @floriscornel/teams-mcp@latest authenticate
```

Check your authentication status:

```bash
npx @floriscornel/teams-mcp@latest check
```

Logout if needed:

```bash
npx @floriscornel/teams-mcp@latest logout
```

### Integrating with Cursor/Claude

This MCP server is designed to work with AI assistants like Claude/Cursor/VS Code through the Model Context Protocol. 

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@floriscornel/teams-mcp@latest"]
    }
  }
}
```

## 🔒 Security

- All authentication is handled through Microsoft's OAuth 2.0 flow
- Tokens are securely stored and automatically refreshed
- No sensitive data is logged or exposed
- Follows Microsoft Graph API security best practices

## 📝 License

MIT License - see LICENSE file for details

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run linting and formatting
5. Submit a pull request

## 📞 Support

For issues and questions:
- Check the existing GitHub issues
- Review Microsoft Graph API documentation
- Ensure proper authentication and permissions are configured 