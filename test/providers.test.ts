import { describe, expect, it, vi } from "vitest";
import { GmailApiProvider } from "../src/providers/gmail-api.js";
import { GmailImapProvider } from "../src/providers/gmail-imap.js";
import { OutlookProvider } from "../src/providers/outlook.js";

const b64 = (s: string) => Buffer.from(s).toString("base64url");

describe("GmailApiProvider", () => {
  it("lists with query, gets full, sends a reply in the thread", async () => {
    const get = vi.fn(async (p: { id: string; format: string }) => ({
      data: { id: p.id, threadId: "T", snippet: "s", internalDate: "1700000000000", payload: { mimeType: "text/plain", body: { data: b64("body") }, headers: [{ name: "Subject", value: "Sub" }, { name: "Message-ID", value: "<orig@x>" }] } },
    }));
    const list = vi.fn(async () => ({ data: { messages: [{ id: "m1" }] } }));
    const send = vi.fn(async (_: unknown) => ({ data: { id: "sent1" } }));
    const gmail = { users: { messages: { list, get, send } } } as never;
    const p = new GmailApiProvider("g", gmail);
    const l = await p.listMessages({ query: "from:bob", limit: 5 });
    expect(list).toHaveBeenCalledWith({ userId: "me", q: "from:bob", labelIds: ["INBOX"], maxResults: 5 });
    expect(l[0]).toMatchObject({ id: "m1", subject: "Sub" });
    expect(l[0]!.body).toBeUndefined();
    expect((await p.getMessage("m1")).body).toBe("body");
    await p.sendMessage({ to: ["t@x.com"], subject: "Re", body: "hello", replyToMessageId: "m1" });
    const arg = send.mock.calls[0]![0] as { requestBody: { raw: string; threadId: string } };
    expect(arg.requestBody.threadId).toBe("T");
    const raw = Buffer.from(arg.requestBody.raw, "base64url").toString();
    expect(raw).toMatch(/In-Reply-To: <orig@x>/);
    expect(raw).toMatch(/To: t@x.com/);
  });
});

describe("OutlookProvider", () => {
  it("calls Graph with bearer token, $search, and sendMail", async () => {
    const calls: { url: string; init: RequestInit }[] = [];
    const f = vi.fn(async (url: string | URL | Request, init?: RequestInit) => {
      calls.push({ url: String(url), init: init ?? {} });
      if (String(url).includes("sendMail")) return new Response(null, { status: 202 });
      if (String(url).includes("/me/messages/")) return Response.json({ id: "x", body: { contentType: "text", content: "full" } });
      return Response.json({ value: [{ id: "1", subject: "old", receivedDateTime: "2024-01-01T00:00:00Z" }, { id: "2", subject: "new", receivedDateTime: "2024-02-01T00:00:00Z" }] });
    });
    const p = new OutlookProvider("o", async () => "TOKEN", f as never);
    const l = await p.listMessages({ query: "invoice", limit: 10 });
    expect(l.map((m) => m.id)).toEqual(["2", "1"]);
    expect(calls[0]!.url).toContain("/me/mailFolders/inbox/messages?");
    expect(decodeURIComponent(calls[0]!.url)).toContain('$search="invoice"');
    expect((calls[0]!.init.headers as Record<string, string>)["Authorization"]).toBe("Bearer TOKEN");
    expect((await p.getMessage("x")).body).toBe("full");
    await p.sendMessage({ to: ["a@b.com"], subject: "S", body: "B" });
    const sent = JSON.parse(calls[2]!.init.body as string);
    expect(sent.message.toRecipients[0].emailAddress.address).toBe("a@b.com");
  });

  it("throws on Graph errors", async () => {
    const p = new OutlookProvider("o", async () => "T", (async () => new Response("denied", { status: 403 })) as never);
    await expect(p.listMessages({ limit: 1 })).rejects.toThrow(/403/);
  });
});

describe("GmailImapProvider", () => {
  const release = vi.fn();
  const search = vi.fn(async () => [1, 2, 3]);
  const fake = {
    connect: vi.fn(async () => undefined),
    logout: vi.fn(async () => undefined),
    getMailboxLock: vi.fn(async () => ({ release })),
    search,
    fetch: async function* () {
      yield { uid: 2, threadId: "t", envelope: { subject: "two", date: new Date("2024-01-01T00:00:00Z") } };
      yield { uid: 3, threadId: "t", envelope: { subject: "three", date: new Date("2024-02-01T00:00:00Z") } };
    },
    fetchOne: vi.fn(async () => ({ uid: 3, threadId: "t", envelope: { subject: "three", messageId: "<m@x>" }, source: Buffer.from("Subject: three\r\n\r\nhello body\r\n") })),
  };
  const sendMail = vi.fn(async () => ({ messageId: "<new@x>" }));
  const p = new GmailImapProvider("g", "me@gmail.com", { connect: () => fake as never, mailer: { sendMail } });

  it("searches with X-GM-RAW and returns newest first", async () => {
    const l = await p.listMessages({ query: "from:bob", limit: 2 });
    expect(search).toHaveBeenCalledWith({ gmraw: "from:bob" }, { uid: true });
    expect(l.map((m) => m.id)).toEqual(["INBOX:3", "INBOX:2"]);
    expect(release).toHaveBeenCalled();
  });

  it("gets a parsed body and sends with In-Reply-To", async () => {
    expect((await p.getMessage("INBOX:3")).body).toBe("hello body\n");
    await p.sendMessage({ to: ["x@y.com"], subject: "s", body: "b", replyToMessageId: "INBOX:3" });
    expect(sendMail).toHaveBeenCalledWith(expect.objectContaining({ from: "me@gmail.com", inReplyTo: "<m@x>" }));
  });
});
