import { Client } from '@microsoft/microsoft-graph-client';
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
import { mapGraphFolder, mapGraphMessage, mapGraphAttachment, buildGraphFilter, resolveWellKnownFolder } from './mapper.js';

export class OutlookAdapter implements EmailProvider {
  readonly providerType: ProviderTypeValue = ProviderType.Outlook;
  private client: InstanceType<typeof Client> | null = null;
  private accountId: string = '';
  private accessToken: string = '';
  private folderIdCache: Map<string, string> = new Map();

  async connect(credentials: AccountCredentials): Promise<void> {
    if (!credentials.oauth) {
      throw new Error('Outlook adapter requires OAuth credentials');
    }
    this.accountId = credentials.id;
    this.accessToken = credentials.oauth.access_token;

    const client = Client.init({
      authProvider: (done) => {
        done(null, this.accessToken);
      },
    });

    // Wrap the api() method to inject `Prefer: IdType="ImmutableId"` on every
    // request.  This makes the Graph API return stable, immutable message IDs
    // that survive moves across folders.  Without this, older Outlook.com
    // messages can have IDs in a shorter format that the API rejects as
    // "Id is malformed" when the message has been moved internally.
    const originalApi = client.api.bind(client);
    client.api = (path: string) => {
      return originalApi(path).header('Prefer', 'IdType="ImmutableId"');
    };

    this.client = client;
  }

  async disconnect(): Promise<void> {
    this.client = null;
    this.accessToken = '';
  }

  async testConnection(): Promise<{ success: boolean; folderCount: number; error?: string }> {
    try {
      const folders = await this.listFolders();
      return { success: true, folderCount: folders.length };
    } catch (error: any) {
      return { success: false, folderCount: 0, error: error.message };
    }
  }

  private ensureClient(): InstanceType<typeof Client> {
    if (!this.client) throw new Error('Not connected');
    return this.client;
  }

  /**
   * Resolves a folder name/display name/ID to a valid Graph API folder reference.
   * Handles well-known names, localized display names, and raw folder IDs.
   */
  async resolveFolder(nameOrId: string): Promise<string> {
    // Check well-known name mappings (handles localized display names)
    const wellKnown = resolveWellKnownFolder(nameOrId);
    if (wellKnown) return wellKnown;

    // Check cache
    const cached = this.folderIdCache.get(nameOrId.toLowerCase());
    if (cached) return cached;

    // Looks like a raw folder ID (long base64 string) — use as-is
    if (nameOrId.length > 40) return nameOrId;

    // Fall back to listing folders and matching by display name
    const client = this.ensureClient();
    const response = await client.api('/me/mailFolders').get();
    const folders = response.value || [];
    for (const folder of folders) {
      // Cache all folders while we're at it
      this.folderIdCache.set(folder.displayName.toLowerCase(), folder.id);
      if (folder.displayName.toLowerCase() === nameOrId.toLowerCase()) {
        return folder.id;
      }
    }

    // Nothing matched — return original value and let Graph API error naturally
    return nameOrId;
  }

  async listFolders(): Promise<Folder[]> {
    const client = this.ensureClient();
    const response = await client.api('/me/mailFolders').get();
    return (response.value || []).map(mapGraphFolder);
  }

  async createFolder(name: string, parentPath?: string): Promise<Folder> {
    const client = this.ensureClient();
    const endpoint = parentPath
      ? `/me/mailFolders/${encodeURIComponent(parentPath)}/childFolders`
      : '/me/mailFolders';
    const result = await client.api(endpoint).post({ displayName: name });
    return mapGraphFolder(result);
  }

  async search(query: SearchQuery): Promise<Email[]> {
    const client = this.ensureClient();
    let endpoint = '/me/messages';
    if (query.folder) {
      const folderId = await this.resolveFolder(query.folder);
      endpoint = `/me/mailFolders/${encodeURIComponent(folderId)}/messages`;
    }

    const { filter, search } = buildGraphFilter(query);
    let request = client.api(endpoint);

    // When body is not needed, use $select to exclude it — dramatically reduces payload
    if (!query.returnBody) {
      request = request.select('id,conversationId,parentFolderId,from,toRecipients,ccRecipients,bccRecipients,subject,receivedDateTime,bodyPreview,hasAttachments,isRead,importance,flag,isDraft,categories,inferenceClassification,webLink');
    }

    if (filter) {
      request = request.filter(filter);
    }
    if (search) {
      request = request.search(search);
    }
    if (query.limit) {
      request = request.top(query.limit);
    }
    if (query.offset) {
      request = request.skip(query.offset);
    }

    // Graph API does not allow $orderBy when $search is used
    if (!search) {
      request = request.orderby('receivedDateTime desc');
    }

    const response = await request.get();
    return (response.value || []).map((msg: any) => mapGraphMessage(msg, this.accountId));
  }

  async getEmail(id: string): Promise<Email> {
    const client = this.ensureClient();
    const message = await client.api(`/me/messages/${encodeURIComponent(id)}`).get();
    const email = mapGraphMessage(message, this.accountId);

    if (message.hasAttachments) {
      const attachments = await client.api(`/me/messages/${encodeURIComponent(id)}/attachments`).get();
      email.attachments = (attachments.value || []).map(mapGraphAttachment);
    }

    return email;
  }

  async getThread(threadId: string): Promise<Thread> {
    const client = this.ensureClient();
    const response = await client
      .api('/me/messages')
      .filter(`conversationId eq '${threadId}'`)
      .orderby('receivedDateTime asc')
      .get();

    const messages: Email[] = (response.value || []).map((msg: any) =>
      mapGraphMessage(msg, this.accountId)
    );

    const participantMap = new Map<string, { name?: string; email: string }>();
    for (const msg of messages) {
      participantMap.set(msg.from.email, msg.from);
      for (const to of msg.to) {
        participantMap.set(to.email, to);
      }
    }

    return {
      id: threadId,
      subject: messages[0]?.subject || '',
      participants: Array.from(participantMap.values()),
      messageCount: messages.length,
      messages,
      lastMessageDate: messages[messages.length - 1]?.date || '',
    };
  }

  async getAttachment(
    emailId: string,
    attachmentId: string
  ): Promise<{ data: Buffer; meta: AttachmentMeta }> {
    const client = this.ensureClient();
    const attachment = await client
      .api(`/me/messages/${encodeURIComponent(emailId)}/attachments/${encodeURIComponent(attachmentId)}`)
      .get();

    return {
      data: Buffer.from(attachment.contentBytes || '', 'base64'),
      meta: mapGraphAttachment(attachment),
    };
  }

  async sendEmail(params: SendEmailParams): Promise<{ id: string; threadId?: string }> {
    const client = this.ensureClient();

    // True in-thread reply: Graph only threads (preserves conversationId,
    // subject, quoted history) via the createReply draft flow. /sendMail
    // always forks a new conversation regardless of inReplyTo. So when this
    // is a reply, create a reply draft off the original message, overlay our
    // body/recipients/attachments, then send the draft.
    if (params.inReplyTo) {
      // createReply requires the Graph item ID. Prefer replyToGraphId (the
      // tool layer sets it to original.id); fall back to inReplyTo only if it
      // wasn't provided (legacy callers).
      const origId = encodeURIComponent(params.replyToGraphId ?? params.inReplyTo);
      const draft = await client
        .api(`/me/messages/${origId}/createReply`)
        .post({});

      const bodyContent = params.body.html || params.body.text || '';
      const contentType = params.body.html ? 'html' : 'text';
      const toGraphRecipient = (c: { name?: string; email: string }) => ({
        emailAddress: { name: c.name, address: c.email },
      });

      const patch: any = {
        body: { contentType, content: bodyContent },
      };
      // Only override recipients if explicitly provided; otherwise keep the
      // recipients Graph pre-populated on the reply draft (original sender).
      if (params.to?.length) patch.toRecipients = params.to.map(toGraphRecipient);
      if (params.cc?.length) patch.ccRecipients = params.cc.map(toGraphRecipient);
      if (params.bcc?.length) patch.bccRecipients = params.bcc.map(toGraphRecipient);

      await client.api(`/me/messages/${encodeURIComponent(draft.id)}`).patch(patch);

      if (params.attachments?.length) {
        for (const att of params.attachments) {
          await client
            .api(`/me/messages/${encodeURIComponent(draft.id)}/attachments`)
            .post({
              '@odata.type': '#microsoft.graph.fileAttachment',
              name: att.filename,
              contentType: att.contentType,
              contentBytes: att.content.toString('base64'),
            });
        }
      }

      await client.api(`/me/messages/${encodeURIComponent(draft.id)}/send`).post({});
      return { id: draft.id };
    }

    // Fresh email (no thread to reply into): plain sendMail.
    const message = this.buildGraphMessage(params);

    const payload: any = { message };
    if (params.attachments?.length) {
      payload.message.attachments = params.attachments.map((att) => ({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: att.filename,
        contentType: att.contentType,
        contentBytes: att.content.toString('base64'),
      }));
    }

    await client.api('/me/sendMail').post(payload);

    return { id: `sent-${Date.now()}` };
  }

  async createDraft(params: SendEmailParams): Promise<{ id: string }> {
    const client = this.ensureClient();

    // True in-thread reply DRAFT: same as the sendEmail reply path (createReply
    // preserves conversationId/subject/quoted-history), but we STOP before the
    // /send call and leave the draft in place. A plain POST to /me/messages
    // forks a NEW conversation regardless of In-Reply-To headers, which is why
    // header-only threading silently failed on Outlook (orphaned draft, 2026-06-01).
    const replyGraphId = params.replyToGraphId ?? params.inReplyTo;
    if (replyGraphId) {
      const origId = encodeURIComponent(replyGraphId);
      const draft = await client
        .api(`/me/messages/${origId}/createReply`)
        .post({});

      const bodyContent = params.body.html || params.body.text || '';
      const contentType = params.body.html ? 'html' : 'text';
      const toGraphRecipient = (c: { name?: string; email: string }) => ({
        emailAddress: { name: c.name, address: c.email },
      });

      const patch: any = {
        body: { contentType, content: bodyContent },
      };
      // Override the Graph-prepopulated recipients only when explicitly given.
      if (params.to?.length) patch.toRecipients = params.to.map(toGraphRecipient);
      if (params.cc?.length) patch.ccRecipients = params.cc.map(toGraphRecipient);
      if (params.bcc?.length) patch.bccRecipients = params.bcc.map(toGraphRecipient);

      await client.api(`/me/messages/${encodeURIComponent(draft.id)}`).patch(patch);

      if (params.attachments?.length) {
        for (const att of params.attachments) {
          await client
            .api(`/me/messages/${encodeURIComponent(draft.id)}/attachments`)
            .post({
              '@odata.type': '#microsoft.graph.fileAttachment',
              name: att.filename,
              contentType: att.contentType,
              contentBytes: att.content.toString('base64'),
            });
        }
      }

      // NO /send — leave it as a draft in the original conversation.
      return { id: draft.id };
    }

    // Fresh (non-reply) draft: plain message create.
    const message = this.buildGraphMessage(params);
    if (params.attachments?.length) {
      message.attachments = params.attachments.map((att) => ({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: att.filename,
        contentType: att.contentType,
        contentBytes: att.content.toString('base64'),
      }));
    }
    const result = await client.api('/me/messages').post(message);
    return { id: result.id };
  }

  async listDrafts(limit?: number, offset?: number): Promise<Email[]> {
    const client = this.ensureClient();
    let request = client.api('/me/mailFolders/drafts/messages');

    if (limit !== undefined) {
      request = request.top(limit);
    }
    if (offset !== undefined) {
      request = request.skip(offset);
    }

    request = request.orderby('receivedDateTime desc');

    const response = await request.get();
    return (response.value || []).map((msg: any) => mapGraphMessage(msg, this.accountId));
  }

  async moveEmail(emailId: string, targetFolder: string, _sourceFolder?: string): Promise<void> {
    const client = this.ensureClient();
    const destinationId = await this.resolveFolder(targetFolder);
    await client.api(`/me/messages/${encodeURIComponent(emailId)}/move`).post({
      destinationId,
    });
  }

  async deleteEmail(emailId: string, permanent?: boolean, _sourceFolder?: string): Promise<void> {
    const client = this.ensureClient();
    if (permanent) {
      await client.api(`/me/messages/${encodeURIComponent(emailId)}`).delete();
    } else {
      await client.api(`/me/messages/${encodeURIComponent(emailId)}/move`).post({
        destinationId: 'deleteditems',
      });
    }
  }

  async markEmail(
    emailId: string,
    flags: { read?: boolean; starred?: boolean; flagged?: boolean },
    _sourceFolder?: string,
  ): Promise<void> {
    const client = this.ensureClient();
    const patch: any = {};

    if (flags.read !== undefined) {
      patch.isRead = flags.read;
    }
    if (flags.starred !== undefined) {
      patch.importance = flags.starred ? 'high' : 'normal';
    }
    if (flags.flagged !== undefined) {
      patch.flag = { flagStatus: flags.flagged ? 'flagged' : 'notFlagged' };
    }

    await client.api(`/me/messages/${encodeURIComponent(emailId)}`).patch(patch);
  }

  private async executeBatch(requests: Array<{ id: string; method: string; url: string; body?: any }>): Promise<Map<string, { status: number; body?: any }>> {
    const client = this.ensureClient();
    const results = new Map<string, { status: number; body?: any }>();

    // Graph API allows max 20 requests per batch
    for (let i = 0; i < requests.length; i += 20) {
      const chunk = requests.slice(i, i + 20);

      // Use sequential numeric IDs for the batch request `id` field instead of
      // raw message IDs. The batch `id` is a correlation value only — it is
      // case-INsensitive and must be unique within the batch. Outlook message
      // IDs are long base64 strings that can collide under case-insensitive
      // comparison, causing "Id is malformed" or duplicate-id errors.
      const indexToId = new Map<string, string>();
      const batchPayload = {
        requests: chunk.map((req, idx) => {
          const batchId = String(idx);
          indexToId.set(batchId, req.id);
          return {
            id: batchId,
            method: req.method,
            url: req.url,
            ...(req.body ? { body: req.body, headers: { 'Content-Type': 'application/json' } } : {}),
          };
        }),
      };

      const response = await client.api('/$batch').post(batchPayload);
      for (const resp of response.responses || []) {
        // Map the sequential batch id back to the original email id
        const originalId = indexToId.get(resp.id) ?? resp.id;
        results.set(originalId, { status: resp.status, body: resp.body });
      }
    }

    return results;
  }

  async batchDelete(emailIds: string[], permanent?: boolean, _sourceFolder?: string): Promise<BatchResult> {
    const result: BatchResult = { succeeded: [], failed: [] };

    if (permanent) {
      const requests = emailIds.map((id) => ({
        id,
        method: 'DELETE',
        url: `/me/messages/${encodeURIComponent(id)}`,
      }));
      const responses = await this.executeBatch(requests);
      for (const [id, resp] of responses) {
        if (resp.status >= 200 && resp.status < 300) {
          result.succeeded.push(id);
        } else {
          result.failed.push({ id, error: resp.body?.error?.message || `HTTP ${resp.status}` });
        }
      }
    } else {
      const requests = emailIds.map((id) => ({
        id,
        method: 'POST',
        url: `/me/messages/${encodeURIComponent(id)}/move`,
        body: { destinationId: 'deleteditems' },
      }));
      const responses = await this.executeBatch(requests);
      for (const [id, resp] of responses) {
        if (resp.status >= 200 && resp.status < 300) {
          result.succeeded.push(id);
        } else {
          result.failed.push({ id, error: resp.body?.error?.message || `HTTP ${resp.status}` });
        }
      }
    }

    return result;
  }

  async batchMove(emailIds: string[], targetFolder: string, _sourceFolder?: string): Promise<BatchResult> {
    const result: BatchResult = { succeeded: [], failed: [] };
    const destinationId = await this.resolveFolder(targetFolder);

    const requests = emailIds.map((id) => ({
      id,
      method: 'POST',
      url: `/me/messages/${encodeURIComponent(id)}/move`,
      body: { destinationId },
    }));

    const responses = await this.executeBatch(requests);
    for (const [id, resp] of responses) {
      if (resp.status >= 200 && resp.status < 300) {
        result.succeeded.push(id);
      } else {
        result.failed.push({ id, error: resp.body?.error?.message || `HTTP ${resp.status}` });
      }
    }

    return result;
  }

  async batchMark(emailIds: string[], flags: { read?: boolean; starred?: boolean; flagged?: boolean }, _sourceFolder?: string): Promise<BatchResult> {
    const result: BatchResult = { succeeded: [], failed: [] };
    const patch: any = {};

    if (flags.read !== undefined) patch.isRead = flags.read;
    if (flags.starred !== undefined) patch.importance = flags.starred ? 'high' : 'normal';
    if (flags.flagged !== undefined) patch.flag = { flagStatus: flags.flagged ? 'flagged' : 'notFlagged' };

    if (Object.keys(patch).length === 0) {
      result.succeeded = [...emailIds];
      return result;
    }

    const requests = emailIds.map((id) => ({
      id,
      method: 'PATCH',
      url: `/me/messages/${encodeURIComponent(id)}`,
      body: patch,
    }));

    const responses = await this.executeBatch(requests);
    for (const [id, resp] of responses) {
      if (resp.status >= 200 && resp.status < 300) {
        result.succeeded.push(id);
      } else {
        result.failed.push({ id, error: resp.body?.error?.message || `HTTP ${resp.status}` });
      }
    }

    return result;
  }

  async getCategories(): Promise<string[]> {
    const client = this.ensureClient();
    const response = await client.api('/me/outlook/masterCategories').get();
    return (response.value || []).map((cat: any) => cat.displayName);
  }

  private buildGraphMessage(params: SendEmailParams): any {
    const toGraphRecipient = (contact: { name?: string; email: string }) => ({
      emailAddress: { name: contact.name, address: contact.email },
    });

    const bodyContent = params.body.html || params.body.text || '';
    const contentType = params.body.html ? 'html' : 'text';

    const message: any = {
      subject: params.subject,
      body: { contentType, content: bodyContent },
      toRecipients: params.to.map(toGraphRecipient),
    };

    if (params.cc?.length) {
      message.ccRecipients = params.cc.map(toGraphRecipient);
    }
    if (params.bcc?.length) {
      message.bccRecipients = params.bcc.map(toGraphRecipient);
    }

    return message;
  }
}
