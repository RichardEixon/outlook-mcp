/**
 * Outlook MCP Server — main entry point.
 *
 * Starts an Express HTTP server that serves:
 *   GET  /sse        — MCP Server-Sent Events endpoint (Claude connects here)
 *   POST /messages   — MCP message endpoint (Claude posts tool calls here)
 *   GET  /auth       — Start Microsoft OAuth login flow
 *   GET  /auth/callback — OAuth redirect URI
 *   GET  /           — Status page
 *   GET  /health     — Health check (used by Replit keep-alive)
 */

import express, { Request, Response } from "express";
import { randomUUID } from "crypto";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";
import dotenv from "dotenv";

import { getAuthorizationUrl, exchangeCodeForTokens, clearTokenCache } from "./auth";
import { calendarTools, handleCalendarTool } from "./tools/calendar";
import { emailTools, handleEmailTool } from "./tools/email";
import { taskTools, handleTaskTool } from "./tools/tasks";

dotenv.config();

const PORT = process.env.PORT || 3000;

// ── MCP Server setup ──────────────────────────────────────────────────────────

const mcpServer = new Server(
  { name: "outlook-mcp-server", version: "1.0.0" },
  { capabilities: { tools: {} } }
);

/** Returns the complete list of all registered tools. */
mcpServer.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [...calendarTools, ...emailTools, ...taskTools],
}));

/** Routes incoming tool-call requests to the appropriate handler. */
mcpServer.setRequestHandler(CallToolRequestSchema, async (request): Promise<any> => {
  const { name, arguments: args = {} } = request.params;

  try {
    if (calendarTools.some((t) => t.name === name)) {
      return await handleCalendarTool(name, args as Record<string, unknown>);
    }
    if (emailTools.some((t) => t.name === name)) {
      return await handleEmailTool(name, args as Record<string, unknown>);
    }
    if (taskTools.some((t) => t.name === name)) {
      return await handleTaskTool(name, args as Record<string, unknown>);
    }
    throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
  } catch (error) {
    if (error instanceof McpError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[MCP] Error in tool "${name}":`, error);
    return {
      content: [{ type: "text", text: `❌ Error: ${message}` }],
      isError: true,
    };
  }
});

// ── Express app ───────────────────────────────────────────────────────────────

const app = express();
app.use(express.json());

/**
 * Active SSE sessions.
 * Key: sessionId (UUID)  Value: SSEServerTransport
 * Supports multiple concurrent Claude connections (e.g. multiple browser tabs).
 */
const activeSessions = new Map<string, SSEServerTransport>();

// ── MCP Endpoints ──────────────────────────────────────────────────────────

/**
 * Claude connects here to establish the SSE stream.
 * We create a new transport per session so multiple tabs work independently.
 */
app.get("/sse", async (req: Request, res: Response) => {
  const sessionId = randomUUID();
  console.log(`[SSE] New connection — session ${sessionId}`);

  // Tell the client to POST messages to /messages?sessionId=<id>
  const transport = new SSEServerTransport(
    `/messages?sessionId=${sessionId}`,
    res
  );
  activeSessions.set(sessionId, transport);

  res.on("close", () => {
    console.log(`[SSE] Session ${sessionId} closed`);
    activeSessions.delete(sessionId);
  });

  await mcpServer.connect(transport);
});

/**
 * Claude posts tool-call messages here.
 * The sessionId query param routes the message to the right SSE transport.
 */
app.post("/messages", async (req: Request, res: Response) => {
  const sessionId = req.query.sessionId as string;
  const transport = activeSessions.get(sessionId);

  if (!transport) {
    res.status(404).json({ error: "Session not found or expired" });
    return;
  }

  await transport.handlePostMessage(req, res);
});

// ── Auth Endpoints ────────────────────────────────────────────────────────────

/** Redirects the user to Microsoft's OAuth login page. */
app.get("/auth", (_req: Request, res: Response) => {
  try {
    const url = getAuthorizationUrl();
    res.redirect(url);
  } catch (error) {
    res.status(500).send(`<pre>Auth setup error: ${error}</pre>`);
  }
});

/** Microsoft redirects here after the user logs in. */
app.get("/auth/callback", async (req: Request, res: Response) => {
  const { code, error, error_description } = req.query as Record<string, string>;

  if (error) {
    res.status(400).send(`
      <h2>❌ Authentication Failed</h2>
      <p><strong>${error}</strong>: ${error_description}</p>
    `);
    return;
  }

  if (!code) {
    res.status(400).send("<h2>❌ No authorization code received.</h2>");
    return;
  }

  try {
    const tokens = await exchangeCodeForTokens(code);

    // Store refresh token in memory so subsequent calls work immediately
    process.env.MICROSOFT_REFRESH_TOKEN = tokens.refreshToken;
    clearTokenCache();

    res.send(`
      <!DOCTYPE html>
      <html>
        <head><title>Outlook MCP — Auth Success</title></head>
        <body style="font-family: system-ui, sans-serif; max-width: 700px; margin: 40px auto; padding: 0 20px;">
          <h2>✅ Authentication Successful!</h2>
          <p>You are now connected to Microsoft Outlook.</p>

          <h3>Important: Save your Refresh Token</h3>
          <p>Copy the token below and add it as a <strong>Replit Secret</strong> named
             <code>MICROSOFT_REFRESH_TOKEN</code> so it persists after restarts:</p>

          <textarea rows="4" style="width:100%; font-family:monospace; font-size:12px; padding:8px;"
            onclick="this.select()">${tokens.refreshToken}</textarea>

          <h3>Next Steps</h3>
          <ol>
            <li>Copy the token above → Replit sidebar → <em>Secrets</em> → add <code>MICROSOFT_REFRESH_TOKEN</code></li>
            <li>Go back to Claude and test a tool, e.g. "show me my calendar events for today"</li>
          </ol>

          <p style="color: green;">The server is now authenticated and ready to use.</p>
        </body>
      </html>
    `);
  } catch (err) {
    res.status(500).send(`<h2>❌ Token exchange failed</h2><pre>${err}</pre>`);
  }
});

// ── Status / Health Endpoints ─────────────────────────────────────────────────

app.get("/health", (_req: Request, res: Response) => {
  res.json({ status: "ok", authenticated: !!process.env.MICROSOFT_REFRESH_TOKEN });
});

app.get("/", (_req: Request, res: Response) => {
  const authenticated = !!process.env.MICROSOFT_REFRESH_TOKEN;
  const sessions = activeSessions.size;

  res.send(`
    <!DOCTYPE html>
    <html>
      <head><title>Outlook MCP Server</title></head>
      <body style="font-family: system-ui, sans-serif; max-width: 600px; margin: 40px auto; padding: 0 20px;">
        <h1>📅 Outlook MCP Server</h1>
        <table style="width:100%; border-collapse:collapse;">
          <tr><td><strong>Status</strong></td><td>🟢 Running</td></tr>
          <tr><td><strong>Authenticated</strong></td><td>${authenticated ? "✅ Yes" : "❌ No — <a href='/auth'>Click here to log in</a>"}</td></tr>
          <tr><td><strong>Active Claude Sessions</strong></td><td>${sessions}</td></tr>
        </table>

        <h2>Endpoints</h2>
        <ul>
          <li><code>GET /sse</code> — Claude MCP connection</li>
          <li><code>POST /messages</code> — Claude MCP messages</li>
          <li><code>GET /auth</code> — ${authenticated ? "Re-authenticate with Microsoft" : "Authenticate with Microsoft"}</li>
          <li><code>GET /health</code> — Health check</li>
        </ul>

        ${!authenticated ? '<p style="color:red;"><strong>Action required:</strong> <a href="/auth">Authenticate with Microsoft</a> before using this MCP.</p>' : ""}
      </body>
    </html>
  `);
});

// ── Start ─────────────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  console.log(`\n🚀 Outlook MCP Server started`);
  console.log(`   URL:           http://localhost:${PORT}`);
  console.log(`   MCP SSE:       http://localhost:${PORT}/sse`);
  console.log(`   Auth:          http://localhost:${PORT}/auth`);
  console.log(`   Authenticated: ${process.env.MICROSOFT_REFRESH_TOKEN ? "✅ Yes" : "❌ No — visit /auth"}\n`);
});
