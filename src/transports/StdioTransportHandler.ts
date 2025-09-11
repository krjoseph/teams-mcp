import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

export class StdioTransportHandler {
  constructor(private server: McpServer) {}

  async connect(): Promise<void> {
    await this.server.connect(new StdioServerTransport());
    console.error('Microsoft Graph MCP Server running on stdio');
  }
}
