import { chmodSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

export interface Account {
  name: string;
  provider: "gmail" | "outlook";
  mode?: "api" | "imap";
  email: string;
  clientId?: string;
  clientSecret?: string;
  tenant?: string;
}

export interface Config {
  accounts: Account[];
  outlookClientId?: string;
  googleClientId?: string;
  googleClientSecret?: string;
}

export function baseDir(): string {
  return process.env["MYEMAILS_DIR"] ?? join(homedir(), ".config", "myemails");
}

export function parseConfig(raw: unknown): Config {
  const cfg = (Array.isArray(raw) ? { accounts: raw } : raw) as Partial<Config> | null;
  if (!cfg || !Array.isArray(cfg.accounts)) throw new Error("accounts.json needs an 'accounts' array");
  const names = new Set<string>();
  for (const a of cfg.accounts) {
    if (!a || typeof a.name !== "string" || !a.name) throw new Error("account without a name");
    if (!/^[A-Za-z0-9._-]+$/.test(a.name)) throw new Error(`account name '${a.name}' may only contain letters, digits, . _ -`);
    if (a.provider !== "gmail" && a.provider !== "outlook") throw new Error(`account '${a.name}': provider must be gmail or outlook`);
    if (typeof a.email !== "string" || !a.email) throw new Error(`account '${a.name}': email is required`);
    if (a.mode !== undefined && a.mode !== "api" && a.mode !== "imap") throw new Error(`account '${a.name}': mode must be api or imap`);
    if (names.has(a.name)) throw new Error(`duplicate account name '${a.name}'`);
    names.add(a.name);
  }
  return cfg as Config;
}

export function loadConfig(): Config {
  const file = join(baseDir(), "accounts.json");
  if (!existsSync(file)) throw new Error(`Config not found: ${file}`);
  return parseConfig(JSON.parse(readFileSync(file, "utf8")));
}

export function findAccount(cfg: Config, name: string): Account {
  const a = cfg.accounts.find((x) => x.name === name);
  if (!a) throw new Error(`Unknown account '${name}'. Known: ${cfg.accounts.map((x) => x.name).join(", ") || "none"}`);
  return a;
}

export function resolveOutlookClientId(cfg: Config, a: Account): string {
  const id = a.clientId ?? cfg.outlookClientId ?? process.env["MYEMAILS_OUTLOOK_CLIENT_ID"];
  if (!id) throw new Error(`No Outlook client ID for '${a.name}'. Set outlookClientId in accounts.json or MYEMAILS_OUTLOOK_CLIENT_ID.`);
  return id;
}

export function resolveGoogleClient(cfg: Config, a: Account): { clientId: string; clientSecret: string } {
  const clientId = a.clientId ?? cfg.googleClientId ?? process.env["MYEMAILS_GOOGLE_CLIENT_ID"];
  const clientSecret = a.clientSecret ?? cfg.googleClientSecret ?? process.env["MYEMAILS_GOOGLE_CLIENT_SECRET"];
  if (!clientId || !clientSecret) throw new Error(`No Google client ID/secret for '${a.name}'. Set googleClientId and googleClientSecret in accounts.json.`);
  return { clientId, clientSecret };
}

function tokenFile(name: string): string {
  return join(baseDir(), "tokens", `${name}.json`);
}

export function loadToken<T = Record<string, unknown>>(name: string): T | undefined {
  const f = tokenFile(name);
  return existsSync(f) ? (JSON.parse(readFileSync(f, "utf8")) as T) : undefined;
}

export function saveToken(name: string, data: unknown): void {
  const dir = join(baseDir(), "tokens");
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  const f = tokenFile(name);
  writeFileSync(f, JSON.stringify(data), { mode: 0o600 });
  chmodSync(f, 0o600);
}

export function hasToken(name: string): boolean {
  return existsSync(tokenFile(name));
}
