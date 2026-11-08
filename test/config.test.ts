import { mkdtempSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it } from "vitest";
import { hasToken, loadToken, parseConfig, resolveGoogleClient, resolveOutlookClientId, saveToken } from "../src/config.js";

const acct = { name: "work", provider: "outlook", email: "a@outlook.com" };

describe("config", () => {
  beforeEach(() => {
    process.env["MYEMAILS_DIR"] = mkdtempSync(join(tmpdir(), "myemails-"));
    delete process.env["MYEMAILS_OUTLOOK_CLIENT_ID"];
  });

  it("accepts an object or a bare array", () => {
    expect(parseConfig({ accounts: [acct] }).accounts).toHaveLength(1);
    expect(parseConfig([acct]).accounts).toHaveLength(1);
  });

  it("rejects bad accounts", () => {
    expect(() => parseConfig({ accounts: [{ ...acct, provider: "yahoo" }] })).toThrow(/provider/);
    expect(() => parseConfig({ accounts: [acct, acct] })).toThrow(/duplicate/);
    expect(() => parseConfig({ accounts: [{ ...acct, name: "../x" }] })).toThrow(/name/);
    expect(() => parseConfig({})).toThrow(/accounts/);
  });

  it("resolves client id: account, then config, then env", () => {
    const cfg = parseConfig({ outlookClientId: "cfg", accounts: [acct, { ...acct, name: "b", clientId: "own" }] });
    expect(resolveOutlookClientId(cfg, cfg.accounts[0]!)).toBe("cfg");
    expect(resolveOutlookClientId(cfg, cfg.accounts[1]!)).toBe("own");
    const none = parseConfig({ accounts: [acct] });
    expect(() => resolveOutlookClientId(none, none.accounts[0]!)).toThrow(/client ID/);
    process.env["MYEMAILS_OUTLOOK_CLIENT_ID"] = "env";
    expect(resolveOutlookClientId(none, none.accounts[0]!)).toBe("env");
  });

  it("requires google client id and secret", () => {
    const cfg = parseConfig({ googleClientId: "i", accounts: [{ name: "g", provider: "gmail", email: "g@gmail.com" }] });
    expect(() => resolveGoogleClient(cfg, cfg.accounts[0]!)).toThrow(/secret/);
  });

  it("stores tokens with mode 0600", () => {
    expect(hasToken("work")).toBe(false);
    saveToken("work", { a: 1 });
    expect(hasToken("work")).toBe(true);
    expect(loadToken("work")).toEqual({ a: 1 });
    const f = join(process.env["MYEMAILS_DIR"]!, "tokens", "work.json");
    expect(statSync(f).mode & 0o777).toBe(0o600);
  });
});
