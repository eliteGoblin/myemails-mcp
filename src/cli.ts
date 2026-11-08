#!/usr/bin/env node
import { createInterface } from "node:readline/promises";
import { findAccount, hasToken, loadConfig, resolveGoogleClient, resolveOutlookClientId } from "./config.js";
import { login as gmailApiLogin } from "./providers/gmail-api.js";
import { login as gmailImapLogin } from "./providers/gmail-imap.js";
import { login as outlookLogin } from "./providers/outlook.js";
import type { LoginIO } from "./providers/types.js";
import { buildProvider } from "./registry.js";

const io: LoginIO = {
  print: (t) => process.stdout.write(`${t}\n`),
  prompt: async (q) => {
    const rl = createInterface({ input: process.stdin, output: process.stdout });
    try { return await rl.question(q); } finally { rl.close(); }
  },
};

async function main(argv: string[]): Promise<void> {
  const [cmd, arg] = argv;
  const cfg = loadConfig();
  if (cmd === "accounts" && (arg === "list" || arg === undefined)) {
    for (const a of cfg.accounts) {
      io.print(`${a.name}\t${a.provider}${a.provider === "gmail" ? `/${a.mode ?? "api"}` : ""}\t${a.email}\t${hasToken(a.name) ? "ok" : "needs login"}`);
    }
  } else if (cmd === "auth" && arg) {
    const a = findAccount(cfg, arg);
    if (a.provider === "outlook") await outlookLogin(a, resolveOutlookClientId(cfg, a), io);
    else if ((a.mode ?? "api") === "imap") await gmailImapLogin(a, io);
    else await gmailApiLogin(a, resolveGoogleClient(cfg, a), io);
    io.print(`Logged in: ${a.name}`);
  } else if (cmd === "test" && arg) {
    const msgs = await buildProvider(cfg, findAccount(cfg, arg)).listMessages({ limit: 5 });
    for (const m of msgs) io.print(`${m.date}\t${m.from}\t${m.subject}`);
    io.print(`${msgs.length} message(s)`);
  } else {
    io.print("Usage: myemails accounts list | auth <account> | test <account>");
    process.exitCode = 2;
  }
}

main(process.argv.slice(2)).catch((e) => {
  process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
