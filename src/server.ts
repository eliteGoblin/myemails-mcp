import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { createHandlers } from "./tools.js";

const account = z.string().describe("Account name from accounts.json");
const addrs = z.array(z.string());

export function createServer(handlers: ReturnType<typeof createHandlers>): McpServer {
  const server = new McpServer({ name: "myemails", version: "0.1.0" });
  const reply = (run: () => Promise<unknown>) => async () => {
    try {
      return { content: [{ type: "text" as const, text: JSON.stringify(await run(), null, 2) }] };
    } catch (e) {
      return { isError: true, content: [{ type: "text" as const, text: e instanceof Error ? e.message : String(e) }] };
    }
  };

  server.registerTool("list_accounts", { description: "List configured email accounts and their login status" },
    reply(() => handlers.list_accounts()));
  server.registerTool("list_messages", {
    description: "List recent messages of one account (no body). query: Gmail search syntax for Gmail, free text for Outlook.",
    inputSchema: { account, query: z.string().optional(), limit: z.number().int().min(1).max(100).optional(), folder: z.string().optional() },
  }, (p) => reply(() => handlers.list_messages(p))());
  server.registerTool("get_message", {
    description: "Get one message with its text body",
    inputSchema: { account, id: z.string() },
  }, (p) => reply(() => handlers.get_message(p))());
  server.registerTool("send_message", {
    description: "Send an email from one account. This sends real mail.",
    inputSchema: { account, to: addrs.min(1), subject: z.string(), body: z.string(), cc: addrs.optional(), bcc: addrs.optional(), replyToMessageId: z.string().optional() },
  }, (p) => reply(() => handlers.send_message(p))());
  server.registerTool("search_all", {
    description: "Search every account and merge results by date, newest first",
    inputSchema: { query: z.string(), limit: z.number().int().min(1).max(100).optional() },
  }, (p) => reply(() => handlers.search_all(p))());
  return server;
}
