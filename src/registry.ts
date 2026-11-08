import { loadToken, resolveGoogleClient, resolveOutlookClientId, type Account, type Config } from "./config.js";
import { makeGmailApi } from "./providers/gmail-api.js";
import { makeGmailImap } from "./providers/gmail-imap.js";
import { makeOutlook } from "./providers/outlook.js";
import type { Provider } from "./providers/types.js";

export function buildProvider(cfg: Config, a: Account): Provider {
  const needLogin = () => new Error(`Account '${a.name}' needs login: run 'myemails auth ${a.name}'`);
  const tok = loadToken<{ refresh_token?: string; appPassword?: string }>(a.name);
  if (a.provider === "outlook") {
    if (!tok) throw needLogin();
    return makeOutlook(a, resolveOutlookClientId(cfg, a));
  }
  if ((a.mode ?? "api") === "imap") {
    if (!tok?.appPassword) throw needLogin();
    return makeGmailImap(a, tok.appPassword);
  }
  if (!tok?.refresh_token) throw needLogin();
  return makeGmailApi(a, resolveGoogleClient(cfg, a), tok.refresh_token);
}
