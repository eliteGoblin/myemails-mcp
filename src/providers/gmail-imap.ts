import { ImapFlow } from "imapflow";
import nodemailer, { type SendMailOptions } from "nodemailer";
import { simpleParser } from "mailparser";
import { saveToken, type Account } from "../config.js";
import { fromImap } from "../normalize.js";
import type { ListOptions, LoginIO, Message, Provider, SendInput } from "./types.js";

type Imap = Pick<ImapFlow, "connect" | "logout" | "getMailboxLock" | "search" | "fetch" | "fetchOne">;
interface Mailer { sendMail(o: SendMailOptions): Promise<{ messageId?: string }> }

export interface ImapDeps {
  connect: () => Imap;
  mailer: Mailer;
}

export function defaultDeps(email: string, password: string): ImapDeps {
  return {
    connect: () => new ImapFlow({ host: "imap.gmail.com", port: 993, secure: true, auth: { user: email, pass: password }, logger: false }),
    mailer: nodemailer.createTransport({ host: "smtp.gmail.com", port: 465, secure: true, auth: { user: email, pass: password } }),
  };
}

export class GmailImapProvider implements Provider {
  constructor(private account: string, private email: string, private deps: ImapDeps) {}

  private async withMailbox<T>(folder: string, fn: (c: Imap) => Promise<T>): Promise<T> {
    const c = this.deps.connect();
    await c.connect();
    try {
      const lock = await c.getMailboxLock(folder);
      try { return await fn(c); } finally { lock.release(); }
    } finally {
      await c.logout();
    }
  }

  async listMessages(o: ListOptions): Promise<Message[]> {
    const folder = o.folder ?? "INBOX";
    return this.withMailbox(folder, async (c) => {
      const found = await c.search(o.query ? { gmraw: o.query } : { all: true }, { uid: true });
      const uids = (found || []).slice(-o.limit);
      if (!uids.length) return [];
      const out: Message[] = [];
      for await (const m of c.fetch(uids, { uid: true, envelope: true, threadId: true }, { uid: true })) {
        out.push(fromImap({ folder, uid: m.uid, threadId: m.threadId, envelope: m.envelope ?? {} }, this.account, false));
      }
      return out.sort((a, b) => b.date.localeCompare(a.date));
    });
  }

  async getMessage(id: string): Promise<Message> {
    const i = id.lastIndexOf(":");
    const folder = id.slice(0, i), uid = id.slice(i + 1);
    return this.withMailbox(folder, async (c) => {
      const m = await c.fetchOne(uid, { uid: true, envelope: true, source: true, threadId: true }, { uid: true });
      if (!m || !m.source) throw new Error(`Message ${id} not found`);
      const p = await simpleParser(m.source);
      return fromImap(
        { folder, uid: m.uid, threadId: m.threadId, envelope: m.envelope ?? {}, text: p.text, html: typeof p.html === "string" ? p.html : undefined },
        this.account,
        true,
      );
    });
  }

  async sendMessage(m: SendInput): Promise<{ id: string }> {
    let inReplyTo: string | undefined;
    if (m.replyToMessageId) {
      const orig = await this.getRaw(m.replyToMessageId);
      inReplyTo = orig;
    }
    const info = await this.deps.mailer.sendMail({
      from: this.email, to: m.to, cc: m.cc, bcc: m.bcc, subject: m.subject, text: m.body,
      inReplyTo, references: inReplyTo,
    });
    return { id: info.messageId ?? "" };
  }

  private async getRaw(id: string): Promise<string | undefined> {
    const i = id.lastIndexOf(":");
    return this.withMailbox(id.slice(0, i), async (c) => {
      const m = await c.fetchOne(id.slice(i + 1), { envelope: true }, { uid: true });
      return m ? (m.envelope?.messageId ?? undefined) : undefined;
    });
  }
}

export function makeGmailImap(account: Account, appPassword: string): GmailImapProvider {
  return new GmailImapProvider(account.name, account.email, defaultDeps(account.email, appPassword));
}

export async function login(account: Account, io: LoginIO): Promise<void> {
  io.print("Create an app password at https://myaccount.google.com/apppasswords (2-Step Verification must be on).");
  const pw = (await io.prompt("App password: ")).replace(/\s+/g, "");
  if (!pw) throw new Error("Empty app password");
  const c = defaultDeps(account.email, pw).connect();
  await c.connect();
  await c.logout();
  saveToken(account.name, { appPassword: pw });
}
