import { z } from 'zod';
import { readFileSync } from 'node:fs';
import { basename, extname } from 'node:path';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { AccountManager } from '../account-manager.js';
import type { SendEmailParams } from '../providers/provider.js';

// Map a file path to a SendEmailParams attachment (read from disk server-side).
const MIME_BY_EXT: Record<string, string> = {
  '.pdf': 'application/pdf',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.txt': 'text/plain',
  '.csv': 'text/csv',
  '.json': 'application/json',
  '.zip': 'application/zip',
  '.stl': 'model/stl',
  '.step': 'application/step',
  '.stp': 'application/step',
  '.3mf': 'application/vnd.ms-package.3dmanufacturing-3dmodel+xml',
};

function pathToAttachment(
  filePath: string,
): { filename: string; content: Buffer; contentType: string } {
  const content = readFileSync(filePath);
  const ext = extname(filePath).toLowerCase();
  return {
    filename: basename(filePath),
    content,
    contentType: MIME_BY_EXT[ext] ?? 'application/octet-stream',
  };
}

const ContactSchema = z.object({
  email: z.string(),
  name: z.string().optional(),
});

const BodySchema = z.object({
  text: z.string().optional(),
  html: z.string().optional(),
});

function jsonResult(data: unknown) {
  return { content: [{ type: 'text' as const, text: JSON.stringify(data) }] };
}

export function registerSendingTools(server: McpServer, accountManager: AccountManager): void {
  // --- email_send ---
  server.tool(
    'email_send',
    'Send a new email',
    {
      accountId: z.string(),
      to: z.array(ContactSchema),
      cc: z.array(ContactSchema).optional(),
      bcc: z.array(ContactSchema).optional(),
      subject: z.string(),
      body: BodySchema,
      attachmentPaths: z
        .array(z.string())
        .optional()
        .describe('Absolute file paths to attach (read server-side)'),
    },
    async (args) => {
      try {
        const provider = await accountManager.getProvider(args.accountId);
        const params: SendEmailParams = {
          to: args.to,
          cc: args.cc,
          bcc: args.bcc,
          subject: args.subject,
          body: args.body,
        };
        if (args.attachmentPaths?.length) {
          params.attachments = args.attachmentPaths.map(pathToAttachment);
        }
        const result = await provider.sendEmail(params);
        return jsonResult(result);
      } catch (error: any) {
        return jsonResult({ error: error.message });
      }
    },
  );

  // --- email_reply ---
  server.tool(
    'email_reply',
    'Reply to an existing email',
    {
      accountId: z.string(),
      emailId: z.string(),
      body: BodySchema,
      replyAll: z.boolean().optional(),
      to: z
        .array(ContactSchema)
        .optional()
        .describe('Explicit To recipients. OVERRIDES the default (which is the original sender). Use when the original message was sent by you, so the reply must go to the real correspondent(s) instead of back to yourself. When provided, this replaces the inherited To entirely.'),
      cc: z
        .array(ContactSchema)
        .optional()
        .describe('Explicit CC recipients on the reply. Use this to add a CC even when replyAll is false. Overrides original-thread CC inheritance.'),
      bcc: z
        .array(ContactSchema)
        .optional()
        .describe('Explicit BCC recipients on the reply.'),
      attachmentPaths: z
        .array(z.string())
        .optional()
        .describe('Absolute file paths to attach to the reply (read server-side)'),
      additionalRecipients: z
        .array(ContactSchema)
        .optional()
        .describe('Extra To recipients to add to the reply (e.g. loop in someone not on the original thread). Deduped by email.'),
    },
    async (args) => {
      try {
        const provider = await accountManager.getProvider(args.accountId);
        const original = await provider.getEmail(args.emailId);

        const messageId = original.headers?.['message-id'] ?? original.id;

        // Build subject with Re: prefix (avoid duplicating)
        const subject = original.subject.startsWith('Re: ')
          ? original.subject
          : `Re: ${original.subject}`;

        // Determine recipients
        let to = [original.from];
        let cc: SendEmailParams['cc'];

        if (args.replyAll) {
          // Include all original "to" recipients
          to = [original.from, ...original.to];
          // Include original CC recipients
          cc = original.cc;
        }

        // Explicit `to` OVERRIDES the inherited default entirely. Needed when the
        // original message was sent by the account owner (Jon): the default To is
        // original.from = Jon, which self-addresses the reply. Providing `to`
        // replaces that with the real correspondent(s). (Bug fix 2026-07-01:
        // email_reply had no `to` param, so replies to a Jon-sent message could
        // only go back to Jon.)
        if (args.to?.length) {
          to = args.to;
        }

        // Loop in extra recipients not on the original thread (e.g. Pete).
        if (args.additionalRecipients?.length) {
          const seen = new Set(to.map((c) => c.email.toLowerCase()));
          for (const r of args.additionalRecipients) {
            if (!seen.has(r.email.toLowerCase())) {
              to.push(r);
              seen.add(r.email.toLowerCase());
            }
          }
        }

        // Explicit cc/bcc OVERRIDE thread-inherited values when provided.
        // This is the bug fix for 2026-05-27: prior code only honored cc via
        // replyAll, so a caller could not add a CC to a thread that had no
        // existing CC. Now `cc: [...]` works as the caller expects.
        if (args.cc !== undefined) {
          cc = args.cc;
        }
        const bcc = args.bcc;

        const params: SendEmailParams = {
          to,
          cc,
          bcc,
          subject,
          body: args.body,
          inReplyTo: messageId,
          references: [messageId],
          // Gmail threads on threadId, not just In-Reply-To/References. Without
          // this, replies opened a brand-new conversation. Harmless for Outlook
          // (which threads via replyToGraphId) and IMAP (header-based).
          threadId: original.threadId,
          // Outlook's createReply needs the Graph item ID, not the RFC
          // Message-ID. Pass it explicitly so the adapter threads reliably even
          // when headers['message-id'] is present (which would otherwise feed a
          // non-Graph ID into createReply). Gmail/IMAP ignore this.
          replyToGraphId: original.id,
        };

        if (args.attachmentPaths?.length) {
          params.attachments = args.attachmentPaths.map(pathToAttachment);
        }

        const result = await provider.sendEmail(params);
        return jsonResult(result);
      } catch (error: any) {
        return jsonResult({ error: error.message });
      }
    },
  );

  // --- email_forward ---
  server.tool(
    'email_forward',
    'Forward an email to new recipients. By default preserves original attachments by fetching their bytes server-side and re-attaching to the outgoing message (prior versions dropped attachments — fixed 2026-05-27).',
    {
      accountId: z.string(),
      emailId: z.string(),
      to: z.array(ContactSchema),
      cc: z.array(ContactSchema).optional(),
      bcc: z.array(ContactSchema).optional(),
      body: BodySchema.optional(),
      attachmentPaths: z
        .array(z.string())
        .optional()
        .describe('Additional absolute file paths to attach (read server-side). Stacks on top of preserved original attachments.'),
      preserveOriginalAttachments: z
        .boolean()
        .optional()
        .describe("Whether to preserve the original message's attachments. Defaults to true."),
    },
    async (args) => {
      try {
        const provider = await accountManager.getProvider(args.accountId);
        const original = await provider.getEmail(args.emailId);

        // Build subject with Fwd: prefix (avoid duplicating)
        const subject = original.subject.startsWith('Fwd: ')
          ? original.subject
          : `Fwd: ${original.subject}`;

        // Build forwarded body
        const forwardHeader = [
          '---------- Forwarded message ----------',
          `From: ${original.from.name ? `${original.from.name} <${original.from.email}>` : original.from.email}`,
          `Date: ${original.date}`,
          `Subject: ${original.subject}`,
          `To: ${original.to.map((c) => (c.name ? `${c.name} <${c.email}>` : c.email)).join(', ')}`,
          '',
        ].join('\n');

        const additionalText = args.body?.text ?? '';
        const originalText = original.body.text ?? '';
        const forwardedText = [
          additionalText,
          additionalText ? '\n' : '',
          forwardHeader,
          originalText,
        ].join('');

        let forwardedHtml: string | undefined;
        if (original.body.html || args.body?.html) {
          const additionalHtml = args.body?.html ?? '';
          const originalHtml = original.body.html ?? '';
          forwardedHtml = [
            additionalHtml,
            '<br/><hr/>',
            `<b>---------- Forwarded message ----------</b><br/>`,
            `From: ${original.from.name ? `${original.from.name} &lt;${original.from.email}&gt;` : original.from.email}<br/>`,
            `Date: ${original.date}<br/>`,
            `Subject: ${original.subject}<br/>`,
            `To: ${original.to.map((c) => (c.name ? `${c.name} &lt;${c.email}&gt;` : c.email)).join(', ')}<br/>`,
            '<br/>',
            originalHtml,
          ].join('');
        }

        const params: SendEmailParams = {
          to: args.to,
          cc: args.cc,
          bcc: args.bcc,
          subject,
          body: { text: forwardedText, html: forwardedHtml },
        };

        // Collect attachments: preserved original + path-based additions
        const attachments: NonNullable<SendEmailParams['attachments']> = [];
        const preserve = args.preserveOriginalAttachments ?? true;
        if (preserve && original.attachments?.length) {
          for (const att of original.attachments) {
            try {
              const fetched = await provider.getAttachment(original.id, att.id);
              attachments.push({
                filename: att.filename,
                content: fetched.data,
                contentType: att.contentType,
              });
            } catch (attErr: any) {
              // Don't fail the whole forward if one attachment can't be fetched.
              // Surface as console warning so operator can see it, but proceed.
              console.warn(
                `email_forward: failed to fetch attachment ${att.id} (${att.filename}): ${attErr.message}`,
              );
            }
          }
        }
        if (args.attachmentPaths?.length) {
          for (const p of args.attachmentPaths) {
            attachments.push(pathToAttachment(p));
          }
        }
        if (attachments.length) {
          params.attachments = attachments;
        }

        const result = await provider.sendEmail(params);
        return jsonResult(result);
      } catch (error: any) {
        return jsonResult({ error: error.message });
      }
    },
  );

  // --- email_draft_create ---
  server.tool(
    'email_draft_create',
    'Create a new email draft. Pass inReplyToEmailId to make it a properly threaded reply (sets In-Reply-To/References headers and Gmail threadId).',
    {
      accountId: z.string(),
      to: z.array(ContactSchema),
      subject: z.string(),
      body: BodySchema,
      inReplyToEmailId: z.string().optional(),
      attachmentPaths: z
        .array(z.string())
        .optional()
        .describe('Absolute file paths to attach to the draft (read server-side)'),
    },
    async (args) => {
      try {
        const provider = await accountManager.getProvider(args.accountId);
        const params: SendEmailParams = {
          to: args.to,
          subject: args.subject,
          body: args.body,
        };
        if (args.attachmentPaths?.length) {
          params.attachments = args.attachmentPaths.map(pathToAttachment);
        }

        if (args.inReplyToEmailId) {
          const original = await provider.getEmail(args.inReplyToEmailId);
          const messageId = original.headers?.['message-id'] ?? original.id;
          params.inReplyTo = messageId;
          params.references = [messageId];
          if (original.threadId) {
            params.threadId = original.threadId;
          }
          // Outlook threads a DRAFT only via the Graph createReply flow, which
          // needs the Graph item ID (not the RFC Message-ID above). Pass it
          // explicitly so createDraft can build a true in-thread reply draft.
          // Gmail/IMAP ignore this and keep using the header-based threading.
          params.replyToGraphId = original.id;
        }

        const result = await provider.createDraft(params);
        return jsonResult(result);
      } catch (error: any) {
        return jsonResult({ error: error.message });
      }
    },
  );

  // --- email_draft_list ---
  server.tool(
    'email_draft_list',
    'List email drafts',
    {
      accountId: z.string(),
      limit: z.number().optional(),
      offset: z.number().optional(),
    },
    async (args) => {
      try {
        const provider = await accountManager.getProvider(args.accountId);
        const drafts = await provider.listDrafts(args.limit, args.offset);
        return jsonResult(drafts);
      } catch (error: any) {
        return jsonResult({ error: error.message });
      }
    },
  );
}
