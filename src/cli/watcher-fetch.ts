/**
 * watcher-fetch — headless mail fetch for the BCC inbox-watcher (Phase B).
 *
 * Runs standalone (no MCP server, no stdio transport): loads the same
 * ~/.email-mcp credential store the server uses, refreshes OAuth silently via
 * MSAL's cache, and prints pure JSON to stdout for a Python caller to parse.
 * Errors go to stderr only; exit code 0 = stdout is valid JSON.
 *
 * IMPORTANT: the credential store's encryption key is derived from
 * hostname+username, so this MUST run as the same interactive user that set
 * the accounts up (<user>) — never as SYSTEM.
 *
 * Usage:
 *   node dist/cli/watcher-fetch.js --account <id> --since <iso8601>
 *       [--folder inbox] [--limit 25] [--with-attachments]
 *   node dist/cli/watcher-fetch.js --account <id> --download <emailId> <attachmentId> <outPath>
 *
 * Output (list mode): JSON array of
 *   { id, threadId, from:{name,email}, subject, date, preview,
 *     hasAttachments, attachments:[{id,filename,contentType,size}] }
 *   (attachments populated only with --with-attachments, via per-message getEmail)
 * Output (download mode): { saved: <outPath>, filename, contentType, size }
 */
import { writeFileSync } from 'fs';
import { AccountManager } from '../account-manager.js';
import type { Email } from '../models/types.js';

function argValue(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

function fail(msg: string): never {
  process.stderr.write(`watcher-fetch: ${msg}\n`);
  process.exit(1);
}

async function main(): Promise<void> {
  const account = argValue('--account') ?? fail('--account <id> is required');

  const mgr = new AccountManager();
  const provider = await mgr.getProvider(account);

  const dlIdx = process.argv.indexOf('--download');
  if (dlIdx >= 0) {
    const [emailId, attachmentId, outPath] = process.argv.slice(dlIdx + 1, dlIdx + 4);
    if (!emailId || !attachmentId || !outPath) {
      fail('--download requires <emailId> <attachmentId> <outPath>');
    }
    const { data, meta } = await provider.getAttachment(emailId, attachmentId);
    writeFileSync(outPath, data);
    process.stdout.write(JSON.stringify({
      saved: outPath,
      filename: meta.filename,
      contentType: meta.contentType,
      size: meta.size,
    }));
    return;
  }

  const since = argValue('--since') ?? fail('--since <iso8601> is required in list mode');
  const folder = argValue('--folder') ?? 'inbox';
  const limit = parseInt(argValue('--limit') ?? '25', 10);
  const withAttachments = process.argv.includes('--with-attachments');

  const emails = await provider.search({ folder, since, limit, returnBody: false });

  const out = [] as Array<Record<string, unknown>>;
  for (const e of emails) {
    let attachments = e.attachments;
    if (withAttachments && e.hasAttachments && attachments.length === 0) {
      // search results carry only the hasAttachments flag; fetch metadata
      const full: Email = await provider.getEmail(e.id);
      attachments = full.attachments;
    }
    out.push({
      id: e.id,
      threadId: e.threadId,
      from: e.from,
      subject: e.subject,
      date: e.date,
      preview: e.snippet ?? '',
      hasAttachments: e.hasAttachments ?? attachments.length > 0,
      attachments: attachments.map((a) => ({
        id: a.id,
        filename: a.filename,
        contentType: a.contentType,
        size: a.size,
      })),
    });
  }
  process.stdout.write(JSON.stringify(out));
}

main().catch((err) => {
  process.stderr.write(`watcher-fetch: ${err?.stack ?? String(err)}\n`);
  process.exit(1);
});
