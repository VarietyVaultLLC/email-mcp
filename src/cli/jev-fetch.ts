/**
 * jev-fetch — read-only headless fetch WITH message bodies, for the Jev shadow test.
 * Sibling of watcher-fetch, which returns previews only. Never sends, moves, marks or
 * labels anything.
 *
 * Same credential rule as watcher-fetch: run as the interactive user who set the accounts
 * up, never SYSTEM.
 *
 * Usage:
 *   node dist/cli/jev-fetch.js --account <id> --since <YYYY-MM-DD> [--folder INBOX] [--limit 200]
 * Output: JSON array of
 *   { id, threadId, folder, from:{name,email}, to:[...], subject, date, labels, draft, text, html, attachments:[filename] }
 */
import { AccountManager } from '../account-manager.js';

function argValue(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

async function main(): Promise<void> {
  const account = argValue('--account');
  const since = argValue('--since');
  if (!account || !since) {
    process.stderr.write('jev-fetch: --account and --since are required\n');
    process.exit(1);
  }
  const folder = argValue('--folder');
  const limit = parseInt(argValue('--limit') ?? '200', 10);

  // Never write the shared credential files (credentials.enc, the MSAL cache): this tool is a
  // read-only observer. Gmail: refresh in memory only (store.save is a no-op here). Outlook:
  // MSAL's silent refresh writes its cache, so an expired Outlook token is skipped (logged,
  // exit 0) instead.
  const mgr = new AccountManager();
  const store = (mgr as any).store;
  store.save = async () => {
    process.stderr.write('jev-fetch: token refreshed in memory only; credentials.enc not written\n');
  };
  const creds = await store.get(account);
  if (!creds) {
    process.stderr.write(`jev-fetch: unknown account ${account}\n`);
    process.exit(1);
  }
  if (creds.provider === 'outlook' && creds.oauth?.expiry) {
    const expiry = new Date(creds.oauth.expiry).getTime();
    if (isNaN(expiry) || expiry <= Date.now() + 5 * 60 * 1000) {
      process.stderr.write('jev-fetch: SKIPPED outlook token expired; not refreshing (would write the MSAL cache)\n');
      process.stdout.write('[]');
      return;
    }
  }
  const provider = await mgr.getProvider(account);
  const emails = await provider.search({ folder, since, limit, returnBody: true });

  process.stdout.write(JSON.stringify(emails.map((e) => ({
    id: e.id,
    threadId: e.threadId,
    folder: e.folder,
    from: e.from,
    to: e.to,
    subject: e.subject,
    date: e.date,
    labels: e.labels ?? [],
    draft: (e.labels ?? []).includes('DRAFT') || (e as any).flags?.draft === true,
    text: e.body?.text ?? '',
    html: e.body?.text ? '' : (e.body?.html ?? ''),
    attachments: e.attachments.map((a) => a.filename),
  }))));
}

main().catch((err) => {
  process.stderr.write(`jev-fetch: ${err?.stack ?? String(err)}\n`);
  process.exit(1);
});
