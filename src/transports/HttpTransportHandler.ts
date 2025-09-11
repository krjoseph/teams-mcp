import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import http from 'http';

export interface HttpTransportConfig {
  port?: number;
  host?: string;
}

export class HttpTransportHandler {
  constructor(
    private server: McpServer,
    private config: HttpTransportConfig = {}
  ) {}

  async connect(): Promise<void> {
    const port = this.config.port ?? parseInt(process.env.PORT || '3000', 10);
    const host = this.config.host ?? '0.0.0.0';

    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
    });

    await this.server.connect(transport);

    const httpServer = http.createServer(async (req, res) => {
      console.log(`Received request: ${req.method} ${req.url}`);

      // Capture original end method to log response
      const originalEnd = res.end;
      res.end = function (chunk?: any, encoding?: any, cb?: any) {
        console.log(
          `Response: ${req.method} ${req.url} - Status: ${res.statusCode}`
        );
        return originalEnd.call(this, chunk, encoding, cb);
      };

      if (req.method === 'GET' && req.url === '/health') {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(
          JSON.stringify({
            status: 'healthy',
            server: 'microsoft-teams-mcp',
            version: '0.0.1',
            timestamp: new Date().toISOString(),
          })
        );
        return;
      }

      try {
        await transport.handleRequest(req, res);
      } catch (error) {
        console.error(error);
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(
            JSON.stringify({
              jsonrpc: '2.0',
              error: {
                code: -32603,
                message: 'Internal server error',
              },
            })
          );
        }
      }
    });

    httpServer.listen(port, host, () => {
      console.log(`Microsoft Teams MCP Server listening on http://${host}:${port}/mcp`);
    });
  }
}
