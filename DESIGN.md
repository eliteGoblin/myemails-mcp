# Design

One MCP server (stdio). Each account in `~/.config/myemails/accounts.json` maps to one provider object. Tools take an `account` name.

## Provider interface (`src/providers/types.ts`)

```ts
interface Provider {
  listMessages(o: { query?: string; limit: number; folder?: string }): Promise<Message[]>; // no body
  getMessage(id: string): Promise<Message>;                                               // with body
  sendMessage(m: SendInput): Promise<{ id: string }>;
}
interface SendInput { to: string[]; subject: string; body: string; cc?: string[]; bcc?: string[]; replyToMessageId?: string }
```

Implementations: `gmail-api.ts` (googleapis), `gmail-imap.ts` (imapflow + nodemailer), `outlook.ts` (Microsoft Graph over fetch). Each takes its client as a constructor argument so tests can mock it. Each also exports `login(account, io)` for the CLI.

## Normalized message (`src/normalize.ts`)

`{ id, threadId, account, from, to[], cc[], date (ISO), subject, snippet, body? }`. `body` is plain text; HTML is converted with `html-to-text`. `list` results omit `body`. IMAP ids are `FOLDER:UID`.

## MCP tools (`src/tools.ts` handlers, `src/server.ts` schemas)

| Tool | Input | Output |
|---|---|---|
| `list_accounts` | none | `[{ name, provider, mode, email, auth: "ok" or "needs login" }]` |
| `list_messages` | `account, query?, limit=20, folder?` | `Message[]` without body |
| `get_message` | `account, id` | `Message` with body |
| `send_message` | `account, to[], subject, body, cc?, bcc?, replyToMessageId?` | `{ id }` |
| `search_all` | `query, limit=20` | `Message[]` merged by date, newest first; per-account errors returned in `errors` |

Query meaning: Gmail API uses Gmail search syntax. Gmail IMAP uses `X-GM-RAW` (same syntax). Outlook uses Graph `$search`. `folder`: Gmail label id (default INBOX), IMAP mailbox, Graph folder name or id (default inbox).

## Config and secrets

`~/.config/myemails/accounts.json`:
`{ "outlookClientId"?, "googleClientId"?, "googleClientSecret"?, "accounts": [{ name, provider, mode?, email, clientId?, clientSecret?, tenant? }] }`. A bare array of accounts is also accepted. Per-account values override the top-level values, then env vars `MYEMAILS_OUTLOOK_CLIENT_ID`, `MYEMAILS_GOOGLE_CLIENT_ID`, `MYEMAILS_GOOGLE_CLIENT_SECRET`. Outlook tenant defaults to `common`. Gmail `mode` defaults to `api`.

Tokens: `~/.config/myemails/tokens/<name>.json`, file mode 0600, directory 0700. `MYEMAILS_DIR` overrides the base directory.

## Auth flows (`myemails auth <account>`)

- Outlook: MSAL device-code flow, public client. Scopes `Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read offline_access` (calendar is requested now so phase 2 needs no re-consent). The MSAL cache is saved to the token file. Later calls use `acquireTokenSilent`.
- Gmail API: loopback server on `127.0.0.1:<random port>`, PKCE, browser opened. Scopes `gmail.modify` and `calendar`. The refresh token is saved.
- Gmail IMAP: prompts for an app password and saves it. Hosts `imap.gmail.com:993`, `smtp.gmail.com:465`.

## File layout

`src/index.ts` stdio entry, `src/server.ts` tool schemas, `src/tools.ts` handlers, `src/cli.ts` CLI, `src/config.ts` config and token store, `src/registry.ts` builds providers from config, `src/normalize.ts`, `src/providers/*`. Tests in `test/`.
