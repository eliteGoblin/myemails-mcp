import { createServer } from "node:http";
import { spawn } from "node:child_process";
import { google, type gmail_v1 } from "googleapis";
import nodemailer from "nodemailer";
import { saveToken, type Account } from "../config.js";
import { fromGmail } from "../normalize.js";
import type { ListOptions, LoginIO, Message, Provider, SendInput } from "./types.js";

export const GMAIL_SCOPES = [
  "https://www.googleapis.com/auth/gmail.modify",
  "https://www.googleapis.com/auth/calendar",
];

const HEADERS = ["From", "To", "Cc", "Date", "Subject"];

export class GmailApiProvider implements Provider {
  constructor(private account: string, private gmail: gmail_v1.Gmail) {}

  async listMessages(o: ListOptions): Promise<Message[]> {
    const res = await this.gmail.users.messages.list({
      userId: "me",
      q: o.query,
      labelIds: [o.folder ?? "INBOX"],
      maxResults: o.limit,
    });
    const ids = (res.data.messages ?? []).map((m) => m.id).filter((x): x is string => !!x);
    const msgs = await Promise.all(
      ids.map((id) =>
        this.gmail.users.messages.get({ userId: "me", id, format: "metadata", metadataHeaders: HEADERS }),
      ),
    );
    return msgs.map((r) => fromGmail(r.data, this.account, false));
  }

  async getMessage(id: string): Promise<Message> {
    const r = await this.gmail.users.messages.get({ userId: "me", id, format: "full" });
    return fromGmail(r.data, this.account, true);
  }

  async sendMessage(m: SendInput): Promise<{ id: string }> {
    let threadId: string | undefined;
    let inReplyTo: string | undefined;
    if (m.replyToMessageId) {
      const o = await this.gmail.users.messages.get({
        userId: "me", id: m.replyToMessageId, format: "metadata", metadataHeaders: ["Message-ID"],
      });
      threadId = o.data.threadId ?? undefined;
      inReplyTo = o.data.payload?.headers?.find((h) => h.name?.toLowerCase() === "message-id")?.value ?? undefined;
    }
    const raw = await buildRaw({ ...m, inReplyTo });
    const r = await this.gmail.users.messages.send({
      userId: "me",
      requestBody: { raw: raw.toString("base64url"), threadId },
    });
    return { id: r.data.id ?? "" };
  }
}

export async function buildRaw(m: SendInput & { inReplyTo?: string; from?: string }): Promise<Buffer> {
  const t = nodemailer.createTransport({ streamTransport: true, buffer: true, newline: "unix" });
  const info = await t.sendMail({
    from: m.from, to: m.to, cc: m.cc, bcc: m.bcc, subject: m.subject, text: m.body,
    inReplyTo: m.inReplyTo, references: m.inReplyTo,
  });
  return info.message as Buffer;
}

export function makeGmailApi(account: Account, client: { clientId: string; clientSecret: string }, refreshToken: string): GmailApiProvider {
  const auth = new google.auth.OAuth2(client.clientId, client.clientSecret);
  auth.setCredentials({ refresh_token: refreshToken });
  return new GmailApiProvider(account.name, google.gmail({ version: "v1", auth }));
}

function openBrowser(url: string): void {
  const cmd = process.platform === "darwin" ? "open" : process.platform === "win32" ? "explorer" : "xdg-open";
  try {
    spawn(cmd, [url], { stdio: "ignore", detached: true }).on("error", () => undefined).unref();
  } catch { /* the URL is printed too */ }
}

export async function login(account: Account, client: { clientId: string; clientSecret: string }, io: LoginIO): Promise<void> {
  const server = createServer();
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  const port = (server.address() as { port: number }).port;
  const oauth = new google.auth.OAuth2(client.clientId, client.clientSecret, `http://127.0.0.1:${port}`);
  const { codeVerifier, codeChallenge } = await oauth.generateCodeVerifierAsync();
  const state = Math.random().toString(36).slice(2);
  const url = oauth.generateAuthUrl({
    access_type: "offline", prompt: "consent", scope: GMAIL_SCOPES, state,
    code_challenge: codeChallenge, code_challenge_method: "S256" as never, login_hint: account.email,
  });
  const code = new Promise<string>((resolve, reject) => {
    server.on("request", (req, res) => {
      const u = new URL(req.url ?? "/", `http://127.0.0.1:${port}`);
      const c = u.searchParams.get("code");
      if (!c || u.searchParams.get("state") !== state) { res.writeHead(400).end("Missing code or bad state"); return; }
      res.end("Login complete. You can close this tab.");
      resolve(c);
    });
    setTimeout(() => reject(new Error("Login timed out after 5 minutes")), 300_000).unref();
  });
  io.print(`Open this URL if the browser does not open:\n${url}`);
  openBrowser(url);
  try {
    const { tokens } = await oauth.getToken({ code: await code, codeVerifier });
    if (!tokens.refresh_token) throw new Error("Google returned no refresh token. Remove the app at myaccount.google.com/permissions and retry.");
    saveToken(account.name, { refresh_token: tokens.refresh_token });
  } finally {
    server.close();
  }
}
