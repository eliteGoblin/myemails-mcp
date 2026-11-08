import { describe, expect, it } from "vitest";
import { fromGmail, fromGraph, fromImap } from "../src/normalize.js";

const b64 = (s: string) => Buffer.from(s).toString("base64url");

describe("normalize", () => {
  it("Gmail payload with plain and html parts", () => {
    const m = fromGmail({
      id: "1", threadId: "t1", snippet: "snip", internalDate: "1700000000000",
      payload: {
        mimeType: "multipart/alternative",
        headers: [{ name: "From", value: "Bob <bob@x.com>" }, { name: "To", value: '"Doe, Jane" <j@x.com>, k@x.com' }, { name: "Subject", value: "Hi" }],
        parts: [{ mimeType: "text/plain", body: { data: b64("plain body") } }, { mimeType: "text/html", body: { data: b64("<p>html</p>") } }],
      },
    }, "g1", true);
    expect(m).toMatchObject({ id: "1", threadId: "t1", account: "g1", from: "Bob <bob@x.com>", subject: "Hi", body: "plain body", date: "2023-11-14T22:13:20.000Z" });
    expect(m.to).toEqual(['"Doe, Jane" <j@x.com>', "k@x.com"]);
  });

  it("Gmail html-only converts to text; list has no body", () => {
    const p = { mimeType: "text/html", body: { data: b64("<p>Hello <b>there</b></p>") }, headers: [] };
    expect(fromGmail({ id: "2", payload: p }, "g", true).body).toBe("Hello there");
    expect(fromGmail({ id: "2", payload: p }, "g", false).body).toBeUndefined();
  });

  it("Graph message", () => {
    const m = fromGraph({
      id: "AAA", conversationId: "c1", subject: "S", bodyPreview: "prev", receivedDateTime: "2024-01-02T03:04:05Z",
      from: { emailAddress: { name: "Ann", address: "ann@o.com" } },
      toRecipients: [{ emailAddress: { address: "me@outlook.com" } }],
      body: { contentType: "html", content: "<div>Hi <i>you</i></div>" },
    }, "o1", true);
    expect(m).toMatchObject({ id: "AAA", threadId: "c1", from: "Ann <ann@o.com>", to: ["me@outlook.com"], cc: [], body: "Hi you", date: "2024-01-02T03:04:05.000Z" });
  });

  it("IMAP envelope", () => {
    const m = fromImap({
      folder: "INBOX", uid: 42, threadId: "99",
      envelope: { date: new Date("2024-05-06T07:08:09Z"), subject: "E", from: [{ name: "Zed", address: "z@x.com" }], to: [{ address: "me@gmail.com" }] },
      text: "line one\n\nline two",
    }, "g2", true);
    expect(m).toMatchObject({ id: "INBOX:42", threadId: "99", from: "Zed <z@x.com>", to: ["me@gmail.com"], snippet: "line one line two", body: "line one\n\nline two" });
  });
});
