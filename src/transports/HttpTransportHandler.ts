import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import http from 'http';
import { randomUUID } from 'crypto';
import { sessionStorage } from '../utils/session.js';
import { LRUCache } from 'lru-cache';
import { GraphService } from '../services/graph.js';
import { registerAuthTools } from '../tools/auth.js';
import { registerChatTools } from '../tools/chats.js';
import { registerSearchTools } from '../tools/search.js';
import { registerTeamsTools } from '../tools/teams.js';
import { registerUsersTools } from '../tools/users.js';

interface CachedTransport {
  transport: StreamableHTTPServerTransport;
  id: string;
  createdAt: number;
}

class TransportCache {
  private cache: LRUCache<string, CachedTransport>;

  constructor() {
    this.cache = new LRUCache<string, CachedTransport>({
      max: 100, // Maximum number of transports to cache
      ttl: 30000, // 30 seconds TTL
      dispose: async (value: CachedTransport) => {
        // Called when transport is evicted from cache
        console.log(`Disposing transport ${value.id} from cache`);
        try {
          await value.transport.close();
        } catch (error) {
          console.error(`Error closing transport ${value.id}:`, error);
        }
      },
      updateAgeOnGet: false, // Don't reset TTL on access
    });
  }

  async getTransport(): Promise<CachedTransport> {
    // Create a new transport for each request
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
    });

    // Create a new server instance for each transport to avoid conflicts
    const server = new McpServer(
      { name: 'Microsoft Teams MCP Server', version: '0.3.3' },
      { capabilities: { tools: {} } }
    );

    // Get the GraphService singleton to register tools
    const graphService = GraphService.getInstance();

    // Register all tools on the new server instance
    registerAuthTools(server, graphService);
    registerUsersTools(server, graphService);
    registerTeamsTools(server, graphService);
    registerChatTools(server, graphService);
    registerSearchTools(server, graphService);

    console.log('Registered all tools on isolated server instance');

    const cachedTransport: CachedTransport = {
      transport,
      id: randomUUID(),
      createdAt: Date.now()
    };

    // Connect the new server instance to transport
    await server.connect(transport);
    
    // Store in cache with TTL
    this.cache.set(cachedTransport.id, cachedTransport);
    
    console.log(`Created and cached transport ${cachedTransport.id} with isolated server`);
    return cachedTransport;
  }

  async destroy() {
    console.log('Destroying transport cache');
    this.cache.clear(); // This will trigger dispose for all cached transports
  }
}

export interface HttpTransportConfig {
  port?: number;
  host?: string;
}

export class HttpTransportHandler {
  private transportCache: TransportCache;

  constructor(
    private config: HttpTransportConfig = {}
  ) {
    this.transportCache = new TransportCache();
  }

  async connect(): Promise<void> {
    const port = this.config.port ?? parseInt(process.env.PORT || '3000', 10);
    const host = this.config.host ?? '0.0.0.0';

    const httpServer = http.createServer(async (req, res) => {
      console.log(`Received request: ${req.method} ${req.url}`);
      
      // Set proper HTTP headers for connection management
      res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
      
      // Capture handler reference for use in closures
      const handler = this;
      
      // Set up request timeout (25 seconds to be safe with Heroku's 30s limit)
      const timeout = setTimeout(() => {
        if (!res.headersSent) {
          console.log(`Request timeout for ${req.method} ${req.url}`);
          res.writeHead(408, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            jsonrpc: '2.0',
            error: {
              code: -32000,
              message: 'Request timeout',
            },
            id: null
          }));
        }
      }, 25000); // 25 second timeout

      // Capture original end method to log response and clear timeout
      const originalEnd = res.end;
      res.end = function(chunk?: any, encoding?: any, cb?: any) {
        clearTimeout(timeout);
        const context = sessionStorage.getStore();
        const sessionPrefix = context?.sessionId ? `[${context.sessionId}] ` : '';
        console.log(`${sessionPrefix}Response: ${req.method} ${req.url} - Status: ${res.statusCode}`);
        return originalEnd.call(this, chunk, encoding, cb);
      };

      // Clean up timeout on request close/error
      req.on('close', () => {
        clearTimeout(timeout);
        console.log(`Request closed: ${req.method} ${req.url}`);
      });
      req.on('error', (error) => {
        clearTimeout(timeout);
        console.log(`Request error: ${req.method} ${req.url}`, error.message);
      });

      if (req.method === 'GET' && req.url === '/health') {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(
          JSON.stringify({
            status: 'healthy',
            server: 'microsoft-teams-mcp',
            version: '0.3.3',
            timestamp: new Date().toISOString(),
          })
        );
        return;
      }

      try {
        const sessionId = (req.headers?.['mcp-session-id'] as string) ?? randomUUID();
        
        // Get a transport from cache (creates new one with TTL)
        console.log(`[${sessionId}] Getting transport from cache`);
        const cachedTransport = await handler.transportCache.getTransport();
        console.log(`[${sessionId}] Using cached transport ${cachedTransport.id}`);
        
        await sessionStorage.run({ sessionId }, async () => {
          console.log(`[${sessionId}] Handling request through cached transport`);
          await cachedTransport.transport.handleRequest(req, res);
          console.log(`[${sessionId}] Transport handleRequest completed`);
        });
      } catch (error) {
        clearTimeout(timeout);
        console.error(`Error handling request ${req.method} ${req.url}:`, error);
        
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            jsonrpc: '2.0',
            error: {
              code: -32603,
              message: 'Internal server error',
            },
            id: null
          }));
        }
      }
    });

    // Configure server settings for better connection handling
    httpServer.keepAliveTimeout = 5000; // 5 seconds
    httpServer.headersTimeout = 6000; // 6 seconds  
    httpServer.timeout = 30000; // 30 seconds total timeout
    httpServer.maxHeadersCount = 100;
    
    httpServer.listen(port, host, () => {
      console.log(`Microsoft Teams MCP Server listening on http://${host}:${port}/mcp`);
    });

    // Handle server errors
    httpServer.on('error', (error) => {
      console.error('HTTP Server error:', error);
    });

    // Handle client errors
    httpServer.on('clientError', (error, socket) => {
      console.error('Client error:', error.message);
      if (!socket.destroyed) {
        socket.end('HTTP/1.1 400 Bad Request\r\n\r\n');
      }
    });
  }

  async destroy() {
    await this.transportCache.destroy();
  }
}
