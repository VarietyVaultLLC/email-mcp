import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import type { gmail_v1 } from 'googleapis';
import type { EmailProvider, SendEmailParams } from '../provider.js';
import type {
  Email,
  Folder,
  Thread,
  SearchQuery,
  AttachmentMeta,
  AccountCredentials,
  ProviderTypeValue,
  BatchResult,
} from '../../models/types.js';
import { ProviderType } from '../../models/types.js';
import { mapGmailLabel, mapGmailMessage, buildGmailQuery } from './mapper.js';

export class GmailAdapter implements EmailProvider {
  readonly providerType: ProviderTypeValue = ProviderType.Gmail;
  private gmail: gmail_v1.Gmail | null = null;
  private accountId: string = '';
  private email: string = '';

  async connect(credentials: AccountCredentials): Promise<void> {
    if (!credentials.oauth) {
      throw new Error('Gmail adapter requires OAuth credentials');
    }
    this.accountId = credentials.id;
    this.email = credentials.email;

    const auth = new OAuth2Client();
    auth.setCredentials({
      access_token: credentials.oauth.access_token,
      refresh_token: credentials.oauth.refresh_token,
      expiry_date: credentials.oauth.expiry
        ? new Date(credentials.oauth.expiry).getTime()
        : undefined,
    });

    this.gmail = google.gmail({ version: 'v1', auth });
  }

  async disconnect(): Promise<void> {
    this.gmail = null;
  }

  async testConnection(): Promise<{ success: boolean; folderCount: number; error?: string }> {
    try {
      const folders = await this.listFolders();
      return { success: true, folderCount: folders.length };
    } catch (error: any) {
      return { success: false, folderCount: 0, error: error.message };
    }
  }

  private ensureConnected(): gmail_v1.Gmail {
    if (!this.gmail) throw new Error('Not connected');
    return this.gmail;
  }

  async listFolders(): Promise<Folder[]> {
    const gmail = this.ensureConnected();
    const res = await gmail.users.labels.list({ userId: 'me' });
    const labels = res.data.labels || [];

    // labels.list doesn't include counts — fetch each with labels.get
    const folders: Folder[] = [];
    const batchSize = 10;
    for (let i = 0; i < labels.length; i += batchSize) {
      const batch = labels.slice(i, i + batchSize);
      const details = await Promise.all(
        batch.map((label) =>
          gmail.users.labels.get({ userId: 'me', id: label.id! })
        )
      );
      for (const detail of details) {
        folders.push(mapGmailLabel(detail.data));
      }
    }
    return folders;
  }

  async createFolder(name: string): Promise<Folder> {
    const gmail = this.ensureConnected();
    const res = await gmail.users.labels.create({
      userId: 'me',
      requestBody: { name },
    });
    return mapGmailLabel(res.data);
  }

  async search(query: SearchQuery): Promise<Email[]> {
    const gmail = this.ensureConnected();
    const q = buildGmailQuery(query);

    const params: any = {
      userId: 'me',
      maxResults: query.limit || 200,
    };

    if (q) params.q = q;
    if (query.folder) params.labelIds = [query.folder];

    const res = await gmail.users.messages.list(params);
    const messages = res.data.messages || [];

    if (messages.length === 0) return [];

    // Fetch full message details in parallel batches of 10
    const emails: Email[] = [];
    const batchSize = 10;
    for (let i = 0; i < messages.length; i += batchSize) {
      const batch = messages.slice(i, i + batchSize);
      const results = await Promise.all(
        batch.map((msg) =>
          gmail.users.messages.get({
            userId: 'me',
            id: msg.id!,
            format: 'full',
          })
        )
      );
      for (const full of results) {
        emails.push(mapGmailMessage(full.data, this.accountId));
      }
    }

    return emails;
  }

  async getEmail(id: string): Promise<Email> {
    const gmail = this.ensureConnected();
    const res = await gmail.users.messages.get({
      userId: 'me',
      id,
      format: 'full',
    });
    return mapGmailMessage(res.data, this.accountId);
  }

  async getThread(threadId: string): Promise<Thread> {
    const gmail = this.ensureConnected();
    const res = await gmail.users.threads.get({
      userId: 'me',
      id: threadId,
      format: 'full',
    });

    const gmailMessages = res.data.messages || [];
    const messages = gmailMessages.map((msg: any) =>
      mapGmailMessage(msg, this.accountId),
    );

    // Collect unique participants
    const participantMap = new Map<string, { name?: string; email: string }>();
    for (const msg of messages) {
      participantMap.set(msg.from.email, msg.from);
      for (const to of msg.to) {
        participantMap.set(to.email, to);
      }
    }

    return {
      id: threadId,
      subject: messages[0]?.subject || '(no subject)',
      participants: Array.from(participantMap.values()),
      messageCount: messages.length,
      messages,
      lastMessageDate: messages[messages.length - 1]?.date || new Date().toISOString(),
    };
  }

  async getAttachment(
    emailId: string,
    attachmentId: string,
  ): Promise<{ data: Buffer; meta: AttachmentMeta }> {
    const gmail = this.ensureConnected();

    // Get the attachment data
    const res = await gmail.users.messages.attachments.get({
      userId: 'me',
      messageId: emailId,
      id: attachmentId,
    });

    // Get the email to find attachment metadata
    const email = await this.getEmail(emailId);
    const attachment = email.attachments.find((a) => a.id === attachmentId);

    return {
      data: Buffer.from(res.data.data || '', 'base64url'),
      meta: attachment || {
        id: attachmentId,
        filename: 'unknown',
        contentType: 'application/octet-stream',
        size: res.data.size || 0,
      },
    };
  }

  async sendEmail(params: SendEmailParams): Promise<{ id: string; threadId?: string }> {
    const gmail = this.ensureConnected();
    const raw = this.buildRfc2822(params);

    const res = await gmail.users.messages.send({
      userId: 'me',
      requestBody: { raw },
    });

    return {
      id: res.data.id || '',
      threadId: res.data.threadId || undefined,
    };
  }

  async createDraft(params: SendEmailParams): Promise<{ id: string }> {
    const gmail = this.ensureConnected();
    const raw = this.buildRfc2822(params);

    const message: { raw: string; threadId?: string } = { raw };
    if (params.threadId) {
      message.threadId = params.threadId;
    }

    const res = await gmail.users.drafts.create({
      userId: 'me',
      requestBody: { message },
    });

    return { id: res.data.id || '' };
  }

  async listDrafts(limit?: number, _offset?: number): Promise<Email[]> {
    const gmail = this.ensureConnected();

    const res = await gmail.users.drafts.list({
      userId: 'me',
      maxResults: limit || 20,
    });

    const drafts = res.data.drafts || [];
    const emails: Email[] = [];

    for (const draft of drafts) {
      const full = await gmail.users.drafts.get({
        userId: 'me',
        id: draft.id!,
        format: 'full',
      });
      if (full.data.message) {
        emails.push(mapGmailMessage(full.data.message, this.accountId));
      }
    }

    return emails;
  }

  async moveEmail(emailId: string, targetFolder: string, _sourceFolder?: string): Promise<void> {
    const gmail = this.ensureConnected();

    // Get current labels to determine source
    const email = await this.getEmail(emailId);
    const currentLabels = email.labels || [];

    // Remove folder-type labels, add target
    const systemFolderLabels = ['INBOX', 'SENT', 'TRASH', 'SPAM', 'DRAFT'];
    const removeLabelIds = currentLabels.filter((l) =>
      systemFolderLabels.includes(l),
    );

    await gmail.users.messages.modify({
      userId: 'me',
      id: emailId,
      requestBody: {
        addLabelIds: [targetFolder],
        removeLabelIds,
      },
    });
  }

  async deleteEmail(emailId: string, permanent?: boolean, _sourceFolder?: string): Promise<void> {
    const gmail = this.ensureConnected();

    if (permanent) {
      await gmail.users.messages.delete({ userId: 'me', id: emailId });
    } else {
      await gmail.users.messages.trash({ userId: 'me', id: emailId });
    }
  }

  async markEmail(
    emailId: string,
    flags: { read?: boolean; starred?: boolean; flagged?: boolean },
    _sourceFolder?: string,
  ): Promise<void> {
    const gmail = this.ensureConnected();
    const addLabelIds: string[] = [];
    const removeLabelIds: string[] = [];

    if (flags.read === true) removeLabelIds.push('UNREAD');
    if (flags.read === false) addLabelIds.push('UNREAD');
    if (flags.starred === true || flags.flagged === true) addLabelIds.push('STARRED');
    if (flags.starred === false || flags.flagged === false) removeLabelIds.push('STARRED');

    if (addLabelIds.length > 0 || removeLabelIds.length > 0) {
      await gmail.users.messages.modify({
        userId: 'me',
        id: emailId,
        requestBody: { addLabelIds, removeLabelIds },
      });
    }
  }

  async batchDelete(emailIds: string[], permanent?: boolean, _sourceFolder?: string): Promise<BatchResult> {
    const gmail = this.ensureConnected();
    const result: BatchResult = { succeeded: [], failed: [] };

    if (permanent) {
      // Gmail batchDelete permanently deletes
      try {
        await gmail.users.messages.batchDelete({
          userId: 'me',
          requestBody: { ids: emailIds },
        });
        result.succeeded = [...emailIds];
      } catch (error: any) {
        // If batch fails, try individually
        for (const id of emailIds) {
          try {
            await gmail.users.messages.delete({ userId: 'me', id });
            result.succeeded.push(id);
          } catch (e: any) {
            result.failed.push({ id, error: e.message });
          }
        }
      }
    } else {
      // Move to trash — no native batch, but batchModify can add TRASH label
      try {
        await gmail.users.messages.batchModify({
          userId: 'me',
          requestBody: {
            ids: emailIds,
            addLabelIds: ['TRASH'],
            removeLabelIds: ['INBOX'],
          },
        });
        result.succeeded = [...emailIds];
      } catch (error: any) {
        for (const id of emailIds) {
          try {
            await gmail.users.messages.trash({ userId: 'me', id });
            result.succeeded.push(id);
          } catch (e: any) {
            result.failed.push({ id, error: e.message });
          }
        }
      }
    }

    return result;
  }

  async batchMove(emailIds: string[], targetFolder: string, _sourceFolder?: string): Promise<BatchResult> {
    const gmail = this.ensureConnected();
    const result: BatchResult = { succeeded: [], failed: [] };

    const systemFolderLabels = ['INBOX', 'SENT', 'TRASH', 'SPAM', 'DRAFT'];
    // Determine which labels to remove (remove all system folder labels)
    const removeLabelIds = systemFolderLabels;

    try {
      await gmail.users.messages.batchModify({
        userId: 'me',
        requestBody: {
          ids: emailIds,
          addLabelIds: [targetFolder],
          removeLabelIds,
        },
      });
      result.succeeded = [...emailIds];
    } catch (error: any) {
      // Fallback to individual moves
      for (const id of emailIds) {
        try {
          await this.moveEmail(id, targetFolder);
          result.succeeded.push(id);
        } catch (e: any) {
          result.failed.push({ id, error: e.message });
        }
      }
    }

    return result;
  }

  async batchMark(emailIds: string[], flags: { read?: boolean; starred?: boolean; flagged?: boolean }, _sourceFolder?: string): Promise<BatchResult> {
    const gmail = this.ensureConnected();
    const result: BatchResult = { succeeded: [], failed: [] };

    const addLabelIds: string[] = [];
    const removeLabelIds: string[] = [];

    if (flags.read === true) removeLabelIds.push('UNREAD');
    if (flags.read === false) addLabelIds.push('UNREAD');
    if (flags.starred === true || flags.flagged === true) addLabelIds.push('STARRED');
    if (flags.starred === false || flags.flagged === false) removeLabelIds.push('STARRED');

    if (addLabelIds.length === 0 && removeLabelIds.length === 0) {
      result.succeeded = [...emailIds];
      return result;
    }

    try {
      await gmail.users.messages.batchModify({
        userId: 'me',
        requestBody: { ids: emailIds, addLabelIds, removeLabelIds },
      });
      result.succeeded = [...emailIds];
    } catch (error: any) {
      for (const id of emailIds) {
        try {
          await this.markEmail(id, flags);
          result.succeeded.push(id);
        } catch (e: any) {
          result.failed.push({ id, error: e.message });
        }
      }
    }

    return result;
  }

  async addLabels(emailId: string, labels: string[]): Promise<void> {
    const gmail = this.ensureConnected();
    await gmail.users.messages.modify({
      userId: 'me',
      id: emailId,
      requestBody: { addLabelIds: labels },
    });
  }

  async removeLabels(emailId: string, labels: string[]): Promise<void> {
    const gmail = this.ensureConnected();
    await gmail.users.messages.modify({
      userId: 'me',
      id: emailId,
      requestBody: { removeLabelIds: labels },
    });
  }

  async listLabels(): Promise<Array<{ id: string; name: string; messageCount: number }>> {
    const gmail = this.ensureConnected();
    const res = await gmail.users.labels.list({ userId: 'me' });
    const labels = res.data.labels || [];

    const results: Array<{ id: string; name: string; messageCount: number }> = [];
    const batchSize = 10;
    for (let i = 0; i < labels.length; i += batchSize) {
      const batch = labels.slice(i, i + batchSize);
      const details = await Promise.all(
        batch.map((label: any) =>
          gmail.users.labels.get({ userId: 'me', id: label.id! })
        )
      );
      for (const detail of details) {
        results.push({
          id: detail.data.id!,
          name: detail.data.name!,
          messageCount: detail.data.messagesTotal ?? 0,
        });
      }
    }
    return results;
  }

  private buildRfc2822(params: SendEmailParams): string {
    const lines: string[] = [];
    lines.push(`From: ${this.email}`);
    lines.push(
      `To: ${params.to.map((c) => (c.name ? `"${c.name}" <${c.email}>` : c.email)).join(', ')}`,
    );
    if (params.cc?.length) {
      lines.push(`Cc: ${params.cc.map((c) => c.email).join(', ')}`);
    }
    if (params.bcc?.length) {
      lines.push(`Bcc: ${params.bcc.map((c) => c.email).join(', ')}`);
    }
    lines.push(`Subject: ${params.subject}`);
    if (params.inReplyTo) {
      lines.push(`In-Reply-To: ${params.inReplyTo}`);
    }
    if (params.references?.length) {
      lines.push(`References: ${params.references.join(' ')}`);
    }
    lines.push('MIME-Version: 1.0');
    lines.push('Content-Type: text/plain; charset=UTF-8');
    lines.push('');
    lines.push(params.body.text || params.body.html || '');

    return Buffer.from(lines.join('\r\n')).toString('base64url');
  }
}
