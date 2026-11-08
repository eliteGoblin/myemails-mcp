import { describe, expect, it, vi } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "../src/server.js";
import { createHandlers } from "../src/tools.js";
import type { Message, Provider } from "../src/providers/types.js";

const msg = (account: string, id: string, date: string): Message => ({ id, threadId: "", account, from: "", to: [], cc: [], date, subject: id, snippet: "" });

function setup() {
  const providers: Record<string, Provider> = {
    a: { listMessages: vi.fn(async () => [msg("a", "a1", "2024-01-01T00:00:00Z"), msg("a", "a2", "2024-03-01T00:00:00Z")]), getMessage: vi.fn(async (id) => msg("a", id, "")), sendMessage: vi.fn(async () => ({ id: "sent" })) },
    b: { listMessages: vi.fn(async () => [msg("b", "b1", "2024-02-01T00:00:00Z")]), getMessage: vi.fn(), sendMessage: vi.fn() },
    c: { listMessages: vi.fn(async () => { throw new Error("boom"); }), getMessage: vi.fn(), sendMessage: vi.fn() },
  };
  const config = { accounts: [
    { name: "a", provider: "gmail" as const, email: "a@gmail.com" },
    { name: "b", provider: "outlook" as const, email: "b@outlook.com" },
    { name: "c", provider: "outlook" as const, email: "c@outlook.com" },
  ] };
  const handlers = createHandlers({ config, getProvider: (x) => providers[x.name]!, hasToken: (n) => n !== "c" });
  return { providers, handlers };
}

describe("tool handlers", () => {
  it("list_accounts reports auth status", async () => {
    const r = await setup().handlers.list_accounts();
    expect(r.map((x) => [x.name, x.mode, x.auth])).toEqual([["a", "api", "ok"], ["b", "graph", "ok"], ["c", "graph", "needs login"]]);
  });

  it("routes to the named account with default limit", async () => {
    const { handlers, providers } = setup();
    await handlers.list_messages({ account: "a", query: "is:unread" });
    expect(providers["a"]!.listMessages).toHaveBeenCalledWith({ query: "is:unread", limit: 20, folder: undefined });
    expect(providers["b"]!.listMessages).not.toHaveBeenCalled();
    expect((await handlers.get_message({ account: "a", id: "x" })).id).toBe("x");
    expect(await handlers.send_message({ account: "a", to: ["t@x.com"], subject: "s", body: "b" })).toEqual({ id: "sent" });
    expect(providers["a"]!.sendMessage).toHaveBeenCalledWith({ to: ["t@x.com"], subject: "s", body: "b" });
  });

  it("rejects unknown accounts", async () => {
    await expect(setup().handlers.list_messages({ account: "zz" })).rejects.toThrow(/Unknown account/);
  });

  it("search_all merges by date and reports per-account errors", async () => {
    const r = await setup().handlers.search_all({ query: "q", limit: 2 });
    expect(r.messages.map((m) => m.id)).toEqual(["a2", "b1"]);
    expect(r.errors).toEqual([{ account: "c", error: "boom" }]);
  });
});

describe("MCP server", () => {
  it("lists 5 tools and calls one over a transport", async () => {
    const [ct, st] = InMemoryTransport.createLinkedPair();
    await createServer(setup().handlers).connect(st);
    const client = new Client({ name: "t", version: "0" });
    await client.connect(ct);
    const tools = (await client.listTools()).tools.map((t) => t.name).sort();
    expect(tools).toEqual(["get_message", "list_accounts", "list_messages", "search_all", "send_message"]);
    const res = await client.callTool({ name: "list_messages", arguments: { account: "zz" } });
    expect(res.isError).toBe(true);
  });
});
