import type {
  Email,
  Folder,
  Thread,
  SearchQuery,
  Contact,
  AttachmentMeta,
  AccountCredentials,
  ProviderTypeValue,
  BatchResult,
} from '../models/types.js';

export interface SendEmailParams {
  to: Contact[];
  cc?: Contact[];
  bcc?: Contact[];
  subject: string;
  body: { text?: string; html?: string };
  attachments?: Array<{ filename: string; content: Buffer; contentType: string }>;
  inReplyTo?: string;
  references?: string[];
  threadId?: string;
  // Graph item ID of the message being replied to (Outlook only). Distinct from
  // inReplyTo (the RFC Message-ID used for header-based threading on Gmail/IMAP).
  // Outlook's createReply flow needs the Graph item ID, NOT the RFC Message-ID,
  // so it's passed separately. Gmail/IMAP ignore this field.
  replyToGraphId?: string;
}

export interface EmailProvider {
  readonly providerType: ProviderTypeValue;

  connect(credentials: AccountCredentials): Promise<void>;
  disconnect(): Promise<void>;
  testConnection(): Promise<{ success: boolean; folderCount: number; error?: string }>;

  listFolders(): Promise<Folder[]>;
  createFolder(name: string, parentPath?: string): Promise<Folder>;

  search(query: SearchQuery): Promise<Email[]>;
  getEmail(id: string): Promise<Email>;
  getThread(threadId: string): Promise<Thread>;
  getAttachment(emailId: string, attachmentId: string): Promise<{ data: Buffer; meta: AttachmentMeta }>;

  sendEmail(params: SendEmailParams): Promise<{ id: string; threadId?: string }>;
  createDraft(params: SendEmailParams): Promise<{ id: string }>;
  listDrafts(limit?: number, offset?: number): Promise<Email[]>;

  moveEmail(emailId: string, targetFolder: string, sourceFolder?: string): Promise<void>;
  deleteEmail(emailId: string, permanent?: boolean, sourceFolder?: string): Promise<void>;
  markEmail(emailId: string, flags: { read?: boolean; starred?: boolean; flagged?: boolean }, sourceFolder?: string): Promise<void>;

  // Batch operations (optional - providers that don't implement fall back to sequential)
  batchDelete?(emailIds: string[], permanent?: boolean, sourceFolder?: string): Promise<BatchResult>;
  batchMove?(emailIds: string[], targetFolder: string, sourceFolder?: string): Promise<BatchResult>;
  batchMark?(emailIds: string[], flags: { read?: boolean; starred?: boolean; flagged?: boolean }, sourceFolder?: string): Promise<BatchResult>;

  // Provider-specific (optional)
  addLabels?(emailId: string, labels: string[]): Promise<void>;
  removeLabels?(emailId: string, labels: string[]): Promise<void>;
  listLabels?(): Promise<Array<{ id: string; name: string; messageCount: number }>>;
  getCategories?(): Promise<string[]>;
}
