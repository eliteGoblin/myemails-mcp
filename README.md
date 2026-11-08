# myemails-mcp

An MCP server (stdio) that gives Claude access to several Gmail and outlook.com accounts at once. Each account has a name. Every tool takes that name.

Tools: `list_accounts`, `list_messages`, `get_message`, `send_message`, `search_all`. See [DESIGN.md](DESIGN.md).

## Install

Requires Node 20 or newer. The package is not on npm yet. Install from GitHub:

```bash
npm i -g github:eliteGoblin/myemails-mcp   # builds on install
myemails accounts list
```

Add it to Claude Code for all projects (syntax from https://code.claude.com/docs/en/mcp: options first, then `--`, then the server command):

```bash
claude mcp add --scope user --transport stdio myemails -- npx -y github:eliteGoblin/myemails-mcp
```

After an npm publish the command is `claude mcp add --scope user --transport stdio myemails -- npx -y myemails-mcp`.

## Config

`~/.config/myemails/accounts.json`:

```json
{
  "outlookClientId": "00000000-0000-0000-0000-000000000000",
  "googleClientId": "xxxx.apps.googleusercontent.com",
  "googleClientSecret": "xxxx",
  "accounts": [
    { "name": "personal", "provider": "gmail", "email": "me@gmail.com" },
    { "name": "old-gmail", "provider": "gmail", "mode": "imap", "email": "old@gmail.com" },
    { "name": "hotmail", "provider": "outlook", "email": "me@outlook.com" }
  ]
}
```

- `mode` for Gmail is `api` (default) or `imap` (app password).
- Per-account `clientId`, `clientSecret`, and `tenant` (Outlook, default `common`) override the top-level values. Env vars `MYEMAILS_OUTLOOK_CLIENT_ID`, `MYEMAILS_GOOGLE_CLIENT_ID`, `MYEMAILS_GOOGLE_CLIENT_SECRET` are the last fallback.
- `MYEMAILS_DIR` changes the config directory.

Then log in once per account and check it:

```bash
myemails auth hotmail     # prints a device code; open the URL and enter it
myemails auth personal    # opens a browser (Gmail API)
myemails auth old-gmail   # asks for the app password (Gmail IMAP)
myemails test hotmail     # lists the 5 newest messages
```

## Outlook setup (Microsoft Entra app)

1. Entra admin center (entra.microsoft.com) > Entra ID > App registrations > New registration.
2. Supported account types: "Any Entra ID Tenant + Personal Microsoft accounts".
3. After creating it, copy the Application (client) ID into `outlookClientId`. One app serves all accounts.
4. API permissions > Add a permission > Microsoft Graph > Delegated: `Mail.Read`, `Mail.Send` (`offline_access` and `User.Read` are added by default).
5. Authentication > Advanced settings > Allow public client flows: Yes.

## Gmail setup, API mode (Google Cloud OAuth client)

1. Google Cloud Console: create a project, then APIs & Services > Library > enable Gmail API.
2. OAuth consent screen: User type External, publishing status Testing. Add every Gmail address you will use as a test user.
3. Credentials > Create credentials > OAuth client ID > Application type Desktop app. Put the client ID and secret in `googleClientId` and `googleClientSecret`.
4. Scopes used: `gmail.readonly` and `gmail.send`.

In Testing mode Google expires the refresh token after 7 days. Run `myemails auth <account>` again when `needs login` returns. Publishing the app to production removes the limit, but Google may then ask for app verification for these sensitive scopes.

## Gmail setup, app password (no Google Cloud project)

Turn on 2-Step Verification, create an app password at https://myaccount.google.com/apppasswords, set `"mode": "imap"` on the account, and run `myemails auth <account>`. Search uses Gmail's `X-GM-RAW`, so Gmail search syntax works.

## Security notes

- Tokens and app passwords are stored in plain JSON under `~/.config/myemails/tokens/` (files mode 0600, directory 0700). Anyone who can read your home directory as you can read your mail. Nothing is stored in this repo.
- With `--scope user`, every Claude Code session on this machine can read and send mail from all configured accounts. `send_message` sends real email. Use Claude Code permission prompts for the `send_message` tool if you want to approve each send.
- Outlook `send_message` with `replyToMessageId` uses the Graph reply endpoint. Graph does not return the id of a sent message, so `send_message` returns an empty id for Outlook.

## Development

```bash
npm ci && npm run build && npm test && npm run lint
```

Tests use mocks only. No network.

## Status

Unit tests and an MCP stdio handshake are verified. Real logins against Gmail and Outlook have not been exercised yet.

MIT license.
