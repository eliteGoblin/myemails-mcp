import { PublicClientApplication, type ICachePlugin, type TokenCacheContext } from "@azure/msal-node";
import { loadToken, saveToken, type Account } from "../config.js";
import { fromGraph, type GraphMessage } from "../normalize.js";
import type { ListOptions, LoginIO, Message, Provider, SendInput } from "./types.js";

export const OUTLOOK_SCOPES = ["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite", "User.Read", "offline_access"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const SELECT = "id,conversationId,subject,bodyPreview,receivedDateTime,from,toRecipients,ccRecipients";

type Fetch = typeof fetch;

export class OutlookProvider implements Provider {
  constructor(
    private account: string,
    private getToken: () => Promise<string>,
    private doFetch: Fetch = fetch,
  ) {}

  private async call(path: string, init: RequestInit = {}): Promise<unknown> {
    const res = await this.doFetch(`${GRAPH}${path}`, {
      ...init,
      headers: {
        Authorization: `Bearer ${await this.getToken()}`,
        "Content-Type": "application/json",
        Prefer: 'outlook.body-content-type="text"',
        ...(init.headers as Record<string, string> | undefined),
      },
    });
    if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
    return res.status === 202 || res.status === 204 ? undefined : res.json();
  }

  async listMessages(o: ListOptions): Promise<Message[]> {
    const p = new URLSearchParams({ $top: String(o.limit), $select: SELECT });
    if (o.query) p.set("$search", `"${o.query.replace(/"/g, "")}"`);
    else p.set("$orderby", "receivedDateTime desc");
    const r = (await this.call(`/me/mailFolders/${encodeURIComponent(o.folder ?? "inbox")}/messages?${p}`)) as { value: GraphMessage[] };
    return r.value
      .map((m) => fromGraph(m, this.account, false))
      .sort((a, b) => b.date.localeCompare(a.date));
  }

  async getMessage(id: string): Promise<Message> {
    const m = (await this.call(`/me/messages/${encodeURIComponent(id)}`)) as GraphMessage;
    return fromGraph(m, this.account, true);
  }

  async sendMessage(m: SendInput): Promise<{ id: string }> {
    const rcpt = (l?: string[]) => (l ?? []).map((address) => ({ emailAddress: { address } }));
    if (m.replyToMessageId) {
      await this.call(`/me/messages/${encodeURIComponent(m.replyToMessageId)}/reply`, {
        method: "POST",
        body: JSON.stringify({
          message: { toRecipients: rcpt(m.to), ccRecipients: rcpt(m.cc), bccRecipients: rcpt(m.bcc) },
          comment: m.body,
        }),
      });
    } else {
      await this.call("/me/sendMail", {
        method: "POST",
        body: JSON.stringify({
          message: {
            subject: m.subject,
            body: { contentType: "Text", content: m.body },
            toRecipients: rcpt(m.to), ccRecipients: rcpt(m.cc), bccRecipients: rcpt(m.bcc),
          },
        }),
      });
    }
    return { id: "" }; // Graph sendMail returns 202 with no id
  }
}

function pca(account: Account, clientId: string): PublicClientApplication {
  const cachePlugin: ICachePlugin = {
    async beforeCacheAccess(ctx: TokenCacheContext) {
      const t = loadToken<{ cache?: string }>(account.name);
      if (t?.cache) ctx.tokenCache.deserialize(t.cache);
    },
    async afterCacheAccess(ctx: TokenCacheContext) {
      if (ctx.cacheHasChanged) saveToken(account.name, { cache: ctx.tokenCache.serialize() });
    },
  };
  return new PublicClientApplication({
    auth: { clientId, authority: `https://login.microsoftonline.com/${account.tenant ?? "common"}` },
    cache: { cachePlugin },
  });
}

export function makeOutlook(account: Account, clientId: string): OutlookProvider {
  const app = pca(account, clientId);
  const getToken = async (): Promise<string> => {
    const accts = await app.getTokenCache().getAllAccounts();
    const a = accts.find((x) => x.username.toLowerCase() === account.email.toLowerCase()) ?? accts[0];
    if (!a) throw new Error(`Account '${account.name}' needs login: run 'myemails auth ${account.name}'`);
    const r = await app.acquireTokenSilent({ account: a, scopes: OUTLOOK_SCOPES });
    return r.accessToken;
  };
  return new OutlookProvider(account.name, getToken);
}

export async function login(account: Account, clientId: string, io: LoginIO): Promise<void> {
  const app = pca(account, clientId);
  await app.acquireTokenByDeviceCode({
    scopes: OUTLOOK_SCOPES,
    deviceCodeCallback: (r) => io.print(r.message),
  });
}
