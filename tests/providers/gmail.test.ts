import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock data defined in vi.hoisted so vi.mock factories can reference it
const mockLabels = [
  { id: 'INBOX', name: 'INBOX', type: 'system', messagesTotal: 100, messagesUnread: 5 },
  { id: 'SENT', name: 'SENT', type: 'system', messagesTotal: 50, messagesUnread: 0 },
  { id: 'DRAFT', name: 'DRAFT', type: 'system', messagesTotal: 3, messagesUnread: 0 },
  { id: 'TRASH', name: 'TRASH', type: 'system', messagesTotal: 10, messagesUnread: 0 },
  { id: 'SPAM', name: 'SPAM', type: 'system', messagesTotal: 8, messagesUnread: 8 },
  { id: 'Label_1', name: 'Work', type: 'user', messagesTotal: 20, messagesUnread: 2 },
];

const mockMessageMetadata = {
  id: 'msg-123',
  threadId: 'thread-456',
  labelIds: ['INBOX', 'UNREAD'],
  snippet: 'Hello, this is a test email...',
  internalDate: '1700000000000',
  payload: {
    headers: [
      { name: 'From', value: 'Alice <alice@example.com>' },
      { name: 'To', value: 'Bob <bob@example.com>' },
      { name: 'Cc', value: 'Carol <carol@example.com>' },
      { name: 'Subject', value: 'Test Email' },
      { name: 'Date', value: 'Tue, 14 Nov 2023 22:13:20 +0000' },
      { name: 'Message-ID', value: '<msg-123@example.com>' },
    ],
    mimeType: 'multipart/alternative',
    parts: [
      {
        mimeType: 'text/plain',
        body: {
          data: Buffer.from('Hello, this is a test email.').toString('base64url'),
          size: 27,
        },
      },
      {
        mimeType: 'text/html',
        body: {
          data: Buffer.from('<p>Hello, this is a test email.</p>').toString('base64url'),
          size: 34,
        },
      },
    ],
  },
};

const mockMessageWithAttachment = {
  id: 'msg-att-1',
  threadId: 'thread-att-1',
  labelIds: ['INBOX'],
  snippet: 'See attached.',
  internalDate: '1700000000000',
  payload: {
    headers: [
      { name: 'From', value: 'alice@example.com' },
      { name: 'To', value: 'bob@example.com' },
      { name: 'Subject', value: 'With Attachment' },
      { name: 'Date', value: 'Tue, 14 Nov 2023 22:13:20 +0000' },
    ],
    mimeType: 'multipart/mixed',
    parts: [
      {
        mimeType: 'text/plain',
        body: { data: Buffer.from('See attached.').toString('base64url'), size: 13 },
      },
      {
        mimeType: 'application/pdf',
        filename: 'report.pdf',
        body: { attachmentId: 'att-001', size: 1024 },
      },
    ],
  },
};

const mockThread = {
  id: 'thread-456',
  messages: [
    mockMessageMetadata,
    {
      ...mockMessageMetadata,
      id: 'msg-789',
      payload: {
        ...mockMessageMetadata.payload,
        headers: [
          { name: 'From', value: 'Bob <bob@example.com>' },
          { name: 'To', value: 'Alice <alice@example.com>' },
          { name: 'Subject', value: 'Re: Test Email' },
          { name: 'Date', value: 'Wed, 15 Nov 2023 10:00:00 +0000' },
        ],
      },
    },
  ],
};

// Mock functions — declared at top level so vi.mock factories can use them
const mockLabelsList = vi.fn();
const mockLabelsCreate = vi.fn();
const mockMessagesList = vi.fn();
const mockMessagesGet = vi.fn();
const mockMessagesSend = vi.fn();
const mockMessagesTrash = vi.fn();
const mockMessagesDelete = vi.fn();
const mockMessagesModify = vi.fn();
const mockAttachmentsGet = vi.fn();
const mockThreadsGet = vi.fn();
const mockDraftsCreate = vi.fn();
const mockDraftsList = vi.fn();
const mockDraftsGet = vi.fn();

// Mock googleapis
vi.mock('googleapis', () => ({
  google: {
    gmail: () => ({
      users: {
        labels: { list: mockLabelsList, create: mockLabelsCreate },
        messages: {
          list: mockMessagesList,
          get: mockMessagesGet,
          send: mockMessagesSend,
          trash: mockMessagesTrash,
          delete: mockMessagesDelete,
          modify: mockMessagesModify,
          attachments: { get: mockAttachmentsGet },
        },
        threads: { get: mockThreadsGet },
        drafts: { create: mockDraftsCreate, list: mockDraftsList, get: mockDraftsGet },
      },
    }),
  },
}));

// Mock google-auth-library
vi.mock('google-auth-library', () => {
  class MockOAuth2Client {
    credentials: Record<string, any> = {};
    setCredentials(creds: any) {
      this.credentials = creds;
    }
  }
  return { OAuth2Client: MockOAuth2Client };
});

import { GmailAdapter } from '../../src/providers/gmail/adapter.js';
import { ProviderType } from '../../src/models/types.js';
import type { AccountCredentials } from '../../src/models/types.js';

const testCredentials: AccountCredentials = {
  id: 'gmail-test-1',
  name: 'Test Gmail',
  provider: 'gmail',
  email: 'test@gmail.com',
  oauth: {
    access_token: 'mock-access-token',
    refresh_token: 'mock-refresh-token',
    expiry: '2030-12-31T00:00:00Z',
  },
};

function resetMocks() {
  mockLabelsList.mockResolvedValue({ data: { labels: mockLabels } });
  mockLabelsCreate.mockResolvedValue({ data: { id: 'new-label', name: 'New', messagesTotal: 0, messagesUnread: 0 } });
  mockMessagesList.mockResolvedValue({
    data: {
      messages: [{ id: 'msg-123', threadId: 'thread-456' }],
      resultSizeEstimate: 1,
    },
  });
  mockMessagesGet.mockResolvedValue({ data: mockMessageMetadata });
  mockMessagesSend.mockResolvedValue({ data: { id: 'sent-1', threadId: 'thread-sent-1' } });
  mockMessagesTrash.mockResolvedValue({ data: {} });
  mockMessagesDelete.mockResolvedValue({ data: {} });
  mockMessagesModify.mockResolvedValue({ data: {} });
  mockAttachmentsGet.mockResolvedValue({ data: { data: Buffer.from('file-content').toString('base64url'), size: 12 } });
  mockThreadsGet.mockResolvedValue({ data: mockThread });
  mockDraftsCreate.mockResolvedValue({ data: { id: 'draft-1', message: { id: 'msg-draft-1' } } });
  mockDraftsList.mockResolvedValue({
    data: { drafts: [{ id: 'draft-1', message: { id: 'msg-draft-1' } }] },
  });
  mockDraftsGet.mockResolvedValue({
    data: { id: 'draft-1', message: mockMessageMetadata },
  });
}

describe('GmailAdapter', () => {
  let adapter: GmailAdapter;

  beforeEach(async () => {
    vi.clearAllMocks();
    resetMocks();

    adapter = new GmailAdapter();
    await adapter.connect(testCredentials);
  });

  it('has correct provider type', () => {
    expect(adapter.providerType).toBe(ProviderType.Gmail);
  });

  describe('connect', () => {
    it('throws when no OAuth credentials', async () => {
      const noOauth = new GmailAdapter();
      await expect(
        noOauth.connect({ ...testCredentials, oauth: undefined }),
      ).rejects.toThrow(/OAuth/);
    });

    it('connects with valid OAuth tokens', async () => {
      const result = await adapter.testConnection();
      expect(result.success).toBe(true);
      expect(result.folderCount).toBeGreaterThan(0);
    });
  });

  describe('listFolders', () => {
    it('maps Gmail labels to Folder[]', async () => {
      const folders = await adapter.listFolders();
      expect(folders.length).toBe(mockLabels.length);

      const inbox = folders.find((f) => f.id === 'INBOX');
      expect(inbox).toBeDefined();
      expect(inbox!.type).toBe('inbox');
      expect(inbox!.totalCount).toBe(100);
      expect(inbox!.unreadCount).toBe(5);

      const sent = folders.find((f) => f.id === 'SENT');
      expect(sent!.type).toBe('sent');

      const trash = folders.find((f) => f.id === 'TRASH');
      expect(trash!.type).toBe('trash');

      const spam = folders.find((f) => f.id === 'SPAM');
      expect(spam!.type).toBe('spam');

      const userLabel = folders.find((f) => f.id === 'Label_1');
      expect(userLabel!.name).toBe('Work');
      expect(userLabel!.type).toBe('other');
    });
  });

  describe('search', () => {
    it('builds Gmail query from SearchQuery', async () => {
      await adapter.search({
        from: 'alice@example.com',
        unreadOnly: true,
        hasAttachment: true,
      });

      expect(mockMessagesList).toHaveBeenCalledWith(
        expect.objectContaining({
          userId: 'me',
          q: expect.stringContaining('from:alice@example.com'),
        }),
      );

      const call = mockMessagesList.mock.calls[0][0];
      expect(call.q).toContain('is:unread');
      expect(call.q).toContain('has:attachment');
    });

    it('applies folder as labelIds', async () => {
      await adapter.search({ folder: 'INBOX' });

      expect(mockMessagesList).toHaveBeenCalledWith(
        expect.objectContaining({
          labelIds: ['INBOX'],
        }),
      );
    });

    it('applies limit as maxResults', async () => {
      await adapter.search({ limit: 10 });

      expect(mockMessagesList).toHaveBeenCalledWith(
        expect.objectContaining({
          maxResults: 10,
        }),
      );
    });

    it('returns mapped Email[]', async () => {
      const emails = await adapter.search({ from: 'alice@example.com' });
      expect(emails).toHaveLength(1);
      expect(emails[0].id).toBe('msg-123');
      expect(emails[0].subject).toBe('Test Email');
    });

    it('returns empty array when no messages found', async () => {
      mockMessagesList.mockResolvedValueOnce({
        data: { messages: undefined, resultSizeEstimate: 0 },
      });
      const emails = await adapter.search({ from: 'nobody@example.com' });
      expect(emails).toHaveLength(0);
    });
  });

  describe('getEmail', () => {
    it('fetches full email by ID', async () => {
      const email = await adapter.getEmail('msg-123');
      expect(email.id).toBe('msg-123');
      expect(email.subject).toBe('Test Email');
      expect(email.from.email).toBe('alice@example.com');
      expect(email.from.name).toBe('Alice');
      expect(email.to[0].email).toBe('bob@example.com');
      expect(email.body.text).toBe('Hello, this is a test email.');
      expect(email.body.html).toContain('<p>Hello');
      expect(email.flags.read).toBe(false); // UNREAD label present
      expect(email.threadId).toBe('thread-456');
    });

    it('calls gmail API with format full', async () => {
      await adapter.getEmail('msg-123');
      expect(mockMessagesGet).toHaveBeenCalledWith({
        userId: 'me',
        id: 'msg-123',
        format: 'full',
      });
    });

    it('maps attachments', async () => {
      mockMessagesGet.mockResolvedValueOnce({
        data: mockMessageWithAttachment,
      });
      const email = await adapter.getEmail('msg-att-1');
      expect(email.attachments).toHaveLength(1);
      expect(email.attachments[0].filename).toBe('report.pdf');
      expect(email.attachments[0].contentType).toBe('application/pdf');
      expect(email.attachments[0].size).toBe(1024);
    });
  });

  describe('getThread', () => {
    it('fetches thread and maps to Thread', async () => {
      const thread = await adapter.getThread('thread-456');
      expect(thread.id).toBe('thread-456');
      expect(thread.messages).toHaveLength(2);
      expect(thread.messageCount).toBe(2);
      expect(thread.subject).toBe('Test Email');
      expect(thread.participants.length).toBeGreaterThan(0);
    });
  });

  describe('sendEmail', () => {
    it('sends email via Gmail API', async () => {
      const result = await adapter.sendEmail({
        to: [{ name: 'Bob', email: 'bob@example.com' }],
        subject: 'Test Send',
        body: { text: 'Hello Bob' },
      });

      expect(result.id).toBe('sent-1');
      expect(result.threadId).toBe('thread-sent-1');
      expect(mockMessagesSend).toHaveBeenCalledWith(
        expect.objectContaining({
          userId: 'me',
          requestBody: expect.objectContaining({ raw: expect.any(String) }),
        }),
      );
    });
  });

  // Regression tests for the 2026-07-14 Gmail send-path bugs. The pre-existing
  // sendEmail test above only asserted `raw: expect.any(String)`, which passed
  // happily while attachments were being dropped on the floor. These decode the
  // actual MIME and assert on its contents.
  describe('MIME construction (regression: silent attachment + html loss)', () => {
    const lastSentMime = (): string => {
      const call = mockMessagesSend.mock.calls.at(-1)?.[0];
      return Buffer.from(call.requestBody.raw, 'base64url').toString('utf8');
    };

    it('includes attachments in the outgoing MIME', async () => {
      await adapter.sendEmail({
        to: [{ email: 'dad@example.com' }],
        subject: 'Reviewed files',
        body: { text: 'Attached.' },
        attachments: [
          {
            filename: 'Details_to_family_REVIEWED.docx',
            content: Buffer.from('PK-fake-docx-bytes'),
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          },
        ],
      });

      const mime = lastSentMime();
      expect(mime).toContain('multipart/mixed');
      expect(mime).toContain('Details_to_family_REVIEWED.docx');
      expect(mime).toContain('Content-Disposition: attachment');
      // The bytes must actually be carried, base64-encoded.
      expect(mime).toContain(Buffer.from('PK-fake-docx-bytes').toString('base64'));
    });

    it('carries multiple attachments', async () => {
      await adapter.sendEmail({
        to: [{ email: 'dad@example.com' }],
        subject: 'Two files',
        body: { text: 'Both attached.' },
        attachments: [
          { filename: 'a.docx', content: Buffer.from('aaa'), contentType: 'application/msword' },
          { filename: 'b.xlsx', content: Buffer.from('bbb'), contentType: 'application/vnd.ms-excel' },
        ],
      });

      const mime = lastSentMime();
      expect(mime).toContain('a.docx');
      expect(mime).toContain('b.xlsx');
    });

    it('sends an HTML body as text/html, not as raw markup in text/plain', async () => {
      await adapter.sendEmail({
        to: [{ email: 'customer@example.com' }],
        subject: 'Branded',
        body: { html: '<div id="sig">Variety Vault 3D</div>' },
      });

      const mime = lastSentMime();
      expect(mime).toContain('text/html');
      // Previously the markup was dumped into a text/plain part, so the
      // recipient saw the literal tags.
      expect(mime).not.toMatch(/Content-Type: text\/plain[\s\S]*?<div id="sig">/);
    });

    it('sends text+html as multipart/alternative', async () => {
      await adapter.sendEmail({
        to: [{ email: 'customer@example.com' }],
        subject: 'Both',
        body: { text: 'plain fallback', html: '<p>rich</p>' },
      });

      const mime = lastSentMime();
      expect(mime).toContain('multipart/alternative');
      expect(mime).toContain('text/plain');
      expect(mime).toContain('text/html');
    });

    it('passes threadId to the API so replies join the parent conversation', async () => {
      await adapter.sendEmail({
        to: [{ email: 'dad@example.com' }],
        subject: 'Re: Family Reunion',
        body: { text: 'reply' },
        inReplyTo: '<parent@mail.gmail.com>',
        references: ['<parent@mail.gmail.com>'],
        threadId: 'thread-parent-1',
      });

      expect(mockMessagesSend).toHaveBeenCalledWith(
        expect.objectContaining({
          requestBody: expect.objectContaining({ threadId: 'thread-parent-1' }),
        }),
      );
      expect(lastSentMime()).toContain('In-Reply-To: <parent@mail.gmail.com>');
    });

    it('omits threadId when not replying', async () => {
      await adapter.sendEmail({
        to: [{ email: 'new@example.com' }],
        subject: 'Fresh',
        body: { text: 'new thread' },
      });

      const call = mockMessagesSend.mock.calls.at(-1)?.[0];
      expect(call.requestBody.threadId).toBeUndefined();
    });

    it('exposes the RFC Message-ID header so email_reply can thread', async () => {
      const email = await adapter.getEmail('msg-123');
      // Without this the reply path fell back to the Gmail internal id and set
      // a bogus In-Reply-To, so every reply opened a new thread.
      expect(email.headers?.['message-id']).toBe('<msg-123@example.com>');
    });
  });

  describe('createDraft', () => {
    it('creates draft via Gmail API', async () => {
      const result = await adapter.createDraft({
        to: [{ email: 'bob@example.com' }],
        subject: 'Draft Test',
        body: { text: 'Draft content' },
      });

      expect(result.id).toBe('draft-1');
      expect(mockDraftsCreate).toHaveBeenCalled();
    });
  });

  describe('listDrafts', () => {
    it('lists drafts', async () => {
      const drafts = await adapter.listDrafts();
      expect(drafts).toHaveLength(1);
    });
  });

  describe('moveEmail', () => {
    it('adds target label and removes source labels', async () => {
      await adapter.moveEmail('msg-123', 'TRASH');

      expect(mockMessagesModify).toHaveBeenCalledWith(
        expect.objectContaining({
          userId: 'me',
          id: 'msg-123',
          requestBody: expect.objectContaining({
            addLabelIds: ['TRASH'],
            removeLabelIds: expect.arrayContaining(['INBOX']),
          }),
        }),
      );
    });
  });

  describe('deleteEmail', () => {
    it('trashes by default', async () => {
      await adapter.deleteEmail('msg-123');
      expect(mockMessagesTrash).toHaveBeenCalledWith({
        userId: 'me',
        id: 'msg-123',
      });
    });

    it('permanently deletes when specified', async () => {
      await adapter.deleteEmail('msg-123', true);
      expect(mockMessagesDelete).toHaveBeenCalledWith({
        userId: 'me',
        id: 'msg-123',
      });
    });
  });

  describe('markEmail', () => {
    it('marks as read by removing UNREAD label', async () => {
      await adapter.markEmail('msg-123', { read: true });
      expect(mockMessagesModify).toHaveBeenCalledWith(
        expect.objectContaining({
          requestBody: expect.objectContaining({
            removeLabelIds: expect.arrayContaining(['UNREAD']),
          }),
        }),
      );
    });

    it('marks as unread by adding UNREAD label', async () => {
      await adapter.markEmail('msg-123', { read: false });
      expect(mockMessagesModify).toHaveBeenCalledWith(
        expect.objectContaining({
          requestBody: expect.objectContaining({
            addLabelIds: expect.arrayContaining(['UNREAD']),
          }),
        }),
      );
    });

    it('stars by adding STARRED label', async () => {
      await adapter.markEmail('msg-123', { starred: true });
      expect(mockMessagesModify).toHaveBeenCalledWith(
        expect.objectContaining({
          requestBody: expect.objectContaining({
            addLabelIds: expect.arrayContaining(['STARRED']),
          }),
        }),
      );
    });
  });

  describe('addLabels', () => {
    it('adds labels via modify', async () => {
      await adapter.addLabels('msg-123', ['Label_1', 'Label_2']);
      expect(mockMessagesModify).toHaveBeenCalledWith({
        userId: 'me',
        id: 'msg-123',
        requestBody: { addLabelIds: ['Label_1', 'Label_2'] },
      });
    });
  });

  describe('removeLabels', () => {
    it('removes labels via modify', async () => {
      await adapter.removeLabels('msg-123', ['Label_1']);
      expect(mockMessagesModify).toHaveBeenCalledWith({
        userId: 'me',
        id: 'msg-123',
        requestBody: { removeLabelIds: ['Label_1'] },
      });
    });
  });

  describe('listLabels', () => {
    it('lists labels with message counts', async () => {
      const labels = await adapter.listLabels();
      expect(labels).toHaveLength(mockLabels.length);
      expect(labels[0]).toHaveProperty('id');
      expect(labels[0]).toHaveProperty('name');
      expect(labels[0]).toHaveProperty('messageCount');
    });
  });
});
