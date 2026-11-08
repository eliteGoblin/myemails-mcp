import { convert } from "html-to-text";
import type { Message } from "./providers/types.js";

export function htmlToText(html: string): string {
  return convert(html, { wordwrap: false, selectors: [{ selector: "img", format: "skip" }, { selector: "a", options: { ignoreHref: true } }] }).trim();
}

const splitAddrs = (s: string): string[] => {
  // split on commas outside quotes and angle brackets
  const out: string[] = [];
  let cur = "", q = false, ang = false;
  for (const c of s) {
    if (c === '"') q = !q;
    else if (c === "<" && !q) ang = true;
    else if (c === ">" && !q) ang = false;
    if (c === "," && !q && !ang) { out.push(cur.trim()); cur = ""; } else cur += c;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
};

const iso = (d: string | number | Date | undefined): string => {
  const t = d === undefined ? NaN : new Date(d).getTime();
  return Number.isNaN(t) ? "" : new Date(t).toISOString();
};

// ---- Gmail API ----
export interface GmailPart {
  mimeType?: string | null;
  body?: { data?: string | null } | null;
  parts?: GmailPart[] | null;
  headers?: { name?: string | null; value?: string | null }[] | null;
}
export interface GmailMessage {
  id?: string | null;
  threadId?: string | null;
  snippet?: string | null;
  internalDate?: string | null;
  payload?: GmailPart | null;
}

const b64 = (s: string): string => Buffer.from(s, "base64url").toString("utf8");

function findPart(p: GmailPart | null | undefined, type: string): string | undefined {
  if (!p) return undefined;
  if (p.mimeType === type && p.body?.data) return b64(p.body.data);
  for (const c of p.parts ?? []) {
    const r = findPart(c, type);
    if (r !== undefined) return r;
  }
  return undefined;
}

export function fromGmail(m: GmailMessage, account: string, withBody: boolean): Message {
  const h = (n: string): string =>
    m.payload?.headers?.find((x) => x.name?.toLowerCase() === n.toLowerCase())?.value ?? "";
  const msg: Message = {
    id: m.id ?? "",
    threadId: m.threadId ?? "",
    account,
    from: h("From"),
    to: splitAddrs(h("To")),
    cc: splitAddrs(h("Cc")),
    date: iso(m.internalDate ? Number(m.internalDate) : h("Date")),
    subject: h("Subject"),
    snippet: m.snippet ?? "",
  };
  if (withBody) {
    const plain = findPart(m.payload, "text/plain");
    const html = findPart(m.payload, "text/html");
    msg.body = plain ?? (html !== undefined ? htmlToText(html) : "");
  }
  return msg;
}

// ---- Microsoft Graph ----
interface GraphAddr { emailAddress?: { name?: string; address?: string } }
export interface GraphMessage {
  id: string;
  conversationId?: string;
  subject?: string;
  bodyPreview?: string;
  receivedDateTime?: string;
  from?: GraphAddr;
  toRecipients?: GraphAddr[];
  ccRecipients?: GraphAddr[];
  body?: { contentType?: string; content?: string };
}

const graphAddr = (a: GraphAddr): string => {
  const { name, address } = a.emailAddress ?? {};
  return name && address ? `${name} <${address}>` : (address ?? name ?? "");
};

export function fromGraph(m: GraphMessage, account: string, withBody: boolean): Message {
  const msg: Message = {
    id: m.id,
    threadId: m.conversationId ?? "",
    account,
    from: m.from ? graphAddr(m.from) : "",
    to: (m.toRecipients ?? []).map(graphAddr),
    cc: (m.ccRecipients ?? []).map(graphAddr),
    date: iso(m.receivedDateTime),
    subject: m.subject ?? "",
    snippet: m.bodyPreview ?? "",
  };
  if (withBody) {
    const c = m.body?.content ?? "";
    msg.body = m.body?.contentType?.toLowerCase() === "html" ? htmlToText(c) : c;
  }
  return msg;
}

// ---- IMAP ----
interface ImapAddr { name?: string; address?: string }
export interface ImapInput {
  folder: string;
  uid: number;
  threadId?: string;
  envelope: { date?: Date | string; subject?: string; from?: ImapAddr[]; to?: ImapAddr[]; cc?: ImapAddr[] };
  text?: string;
  html?: string;
}

const imapAddr = (a: ImapAddr): string => (a.name && a.address ? `${a.name} <${a.address}>` : (a.address ?? a.name ?? ""));

export function fromImap(i: ImapInput, account: string, withBody: boolean): Message {
  const e = i.envelope;
  const text = i.text ?? (i.html ? htmlToText(i.html) : "");
  const msg: Message = {
    id: `${i.folder}:${i.uid}`,
    threadId: i.threadId ?? "",
    account,
    from: (e.from ?? []).map(imapAddr).join(", "),
    to: (e.to ?? []).map(imapAddr),
    cc: (e.cc ?? []).map(imapAddr),
    date: iso(e.date),
    subject: e.subject ?? "",
    snippet: text.replace(/\s+/g, " ").trim().slice(0, 200),
  };
  if (withBody) msg.body = text;
  return msg;
}
