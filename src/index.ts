#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { hasToken, loadConfig } from "./config.js";
import { buildProvider } from "./registry.js";
import { createServer } from "./server.js";
import { createHandlers } from "./tools.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const handlers = createHandlers({ config, getProvider: (a) => buildProvider(config, a), hasToken });
  await createServer(handlers).connect(new StdioServerTransport());
}

main().catch((e) => {
  process.stderr.write(`myemails-mcp: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
