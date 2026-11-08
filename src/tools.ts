import type { Account, Config } from "./config.js";
import type { Message, Provider, SendInput } from "./providers/types.js";

export interface Deps {
  config: Config;
  getProvider: (account: Account) => Provider;
  hasToken: (name: string) => boolean;
}

const find = (d: Deps, name: string): Account => {
  const a = d.config.accounts.find((x) => x.name === name);
  if (!a) throw new Error(`Unknown account '${name}'. Known: ${d.config.accounts.map((x) => x.name).join(", ") || "none"}`);
  return a;
};

export function createHandlers(d: Deps) {
  return {
    list_accounts: async () =>
      d.config.accounts.map((a) => ({
        name: a.name,
        provider: a.provider,
        mode: a.provider === "gmail" ? (a.mode ?? "api") : "graph",
        email: a.email,
        auth: d.hasToken(a.name) ? "ok" : "needs login",
      })),

    list_messages: async (p: { account: string; query?: string; limit?: number; folder?: string }) =>
      d.getProvider(find(d, p.account)).listMessages({ query: p.query, limit: p.limit ?? 20, folder: p.folder }),

    get_message: async (p: { account: string; id: string }) => d.getProvider(find(d, p.account)).getMessage(p.id),

    send_message: async (p: SendInput & { account: string }) => {
      const { account, ...input } = p;
      return d.getProvider(find(d, account)).sendMessage(input);
    },

    search_all: async (p: { query: string; limit?: number }) => {
      const limit = p.limit ?? 20;
      const errors: { account: string; error: string }[] = [];
      const results = await Promise.all(
        d.config.accounts.map(async (a): Promise<Message[]> => {
          try {
            return await d.getProvider(a).listMessages({ query: p.query, limit });
          } catch (e) {
            errors.push({ account: a.name, error: e instanceof Error ? e.message : String(e) });
            return [];
          }
        }),
      );
      const messages = results.flat().sort((a, b) => b.date.localeCompare(a.date)).slice(0, limit);
      return { messages, errors };
    },
  };
}
