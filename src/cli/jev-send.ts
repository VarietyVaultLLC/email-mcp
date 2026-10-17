/**
 * jev-send — headless SELF-send for Jev digests (part-leads daily top 10). The recipient is
 * always the sending account's own address; there is no --to option, so this tool cannot
 * mail anyone else. Gmail only.
 *
 * Same credential rule as jev-fetch: run as the interactive user who set the accounts up,
 * never SYSTEM, and never write the shared credential files (tokens refresh in memory only).
 *
 * Usage:
 *   node dist/cli/jev-send.js --account <id> --subject <text> --body-file <path.txt> [--html-file <path.html>]
 * Output: JSON { id, to }
 */
import fs from 'node:fs';
import { AccountManager } from '../account-manager.js';

function argValue(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

async function main(): Promise<void> {
  const account = argValue('--account');
  const subject = argValue('--subject');
  const bodyFile = argValue('--body-file');
  const htmlFile = argValue('--html-file');
  if (!account || !subject || !bodyFile) {
    process.stderr.write('jev-send: --account, --subject and --body-file are required\n');
    process.exit(1);
  }

  const mgr = new AccountManager();
  const store = (mgr as any).store;
  store.save = async () => {
    process.stderr.write('jev-send: token refreshed in memory only; credentials.enc not written\n');
  };
  const creds = await store.get(account);
  if (!creds) {
    process.stderr.write(`jev-send: unknown account ${account}\n`);
    process.exit(1);
  }
  if (creds.provider !== 'gmail') {
    process.stderr.write(`jev-send: Gmail accounts only (got ${creds.provider})\n`);
    process.exit(1);
  }

  const provider = await mgr.getProvider(account);
  const res = await provider.sendEmail({
    to: [{ email: creds.email }],
    subject,
    body: {
      text: fs.readFileSync(bodyFile, 'utf8'),
      ...(htmlFile ? { html: fs.readFileSync(htmlFile, 'utf8') } : {}),
    },
  });
  process.stdout.write(JSON.stringify({ id: res.id, to: creds.email }));
}

main().catch((err) => {
  process.stderr.write(`jev-send: ${err instanceof Error ? err.message : String(err)}\n`);
  process.exit(1);
});
