export interface Message {
  id: string;
  threadId: string;
  account: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  subject: string;
  snippet: string;
  body?: string;
}

export interface ListOptions {
  query?: string;
  limit: number;
  folder?: string;
}

export interface SendInput {
  to: string[];
  subject: string;
  body: string;
  cc?: string[];
  bcc?: string[];
  replyToMessageId?: string;
}

export interface Provider {
  listMessages(o: ListOptions): Promise<Message[]>;
  getMessage(id: string): Promise<Message>;
  sendMessage(m: SendInput): Promise<{ id: string }>;
}

export interface LoginIO {
  print(text: string): void;
  prompt(question: string): Promise<string>;
}
