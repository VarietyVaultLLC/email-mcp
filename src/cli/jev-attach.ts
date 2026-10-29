/**
 * jev-attach — read-only download of a quote request's 3D/drawing attachments, for the Jev quoting
 * benchmark (VarietyVaultLLC/vv-jev-lab#28). Sibling of jev-fetch. Never sends, moves, marks, labels
 * or deletes mail: it only reads the message (GET) and its attachments (GET).
 *
 * Agreed with guac-bcc 2026-10-29 (7 conditions):
 *   1. Credentials as jev-fetch: refresh in memory only; an expired Outlook token is skipped (exit 0).
 *      Run as the interactive user jonca, never SYSTEM.
 *   2. Filenames come from outsiders: basename only; reject separators, "..", drive letters, control
 *      characters; write under QUOTES\<message-id>\<name>; re-check the resolved path; never overwrite.
 *   3. Allow-list of extensions, 100 MB cap per file, never opened/run/unzipped, .zip refused.
 *   5. No mail operation other than get and attachment fetch.
 *   7. Each download is logged (account, message id, filename, bytes, sha256), never a message body.
 *
 * Usage:
 *   node dist/cli/jev-attach.js --account <id> --message <messageId> [--name <filename>]
 * Output: JSON array of { name, saved?, bytes?, sha256?, skipped? }
 * Build (never touch build.mjs or dist/index.js):
 *   npx esbuild src/cli/jev-attach.ts --bundle --platform=node --target=node20 --format=esm
 *     --outdir=dist --outbase=src --sourcemap --packages=external
 */
import { createHash } from 'node:crypto';
import { appendFileSync, existsSync, mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { AccountManager } from '../account-manager.js';

const QUOTES = 'C:\\Users\\jonca\\Documents\\jev-data\\quotes';
const LOG = path.join(QUOTES, 'downloads.jsonl');
const ALLOWED = new Set(['.stl', '.step', '.stp', '.3mf', '.obj', '.iges', '.igs', '.x_t', '.sldprt', '.pdf']);
const MAX_BYTES = 100 * 1024 * 1024;

function argValue(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

/** Why a sender-supplied filename is refused, or null when it is safe to use as-is. */
export function badName(name: string): string | null {
  if (!name || name.length > 200) return 'empty or too long';
  if (/[\\/]/.test(name)) return 'path separator';
  if (name.includes('..')) return 'dot-dot';
  if (/^[a-zA-Z]:/.test(name) || name.includes(':')) return 'drive letter or colon';
  if (/[\x00-\x1f\x7f]/.test(name)) return 'control character';
  if (path.basename(name) !== name) return 'not a plain file name';
  const ext = path.extname(name).toLowerCase();
  if (ext === '.zip') return 'zip refused';
  if (!ALLOWED.has(ext)) return `type ${ext || '(none)'} not allowed`;
  return null;
}

/** A path under dir that does not exist yet: name, then name-2, name-3 ... (never overwrite). */
function freePath(dir: string, name: string): string {
  const ext = path.extname(name);
  const stem = name.slice(0, name.length - ext.length);
  for (let i = 1; ; i++) {
    const p = path.join(dir, i === 1 ? name : `${stem}-${i}${ext}`);
    if (!existsSync(p)) return p;
  }
}

async function main(): Promise<void> {
  const account = argValue('--account');
  const message = argValue('--message');
  const only = argValue('--name');
  if (!account || !message) {
    process.stderr.write('jev-attach: --account and --message are required\n');
    process.exit(1);
  }

  // Same credential rule as jev-fetch: never write credentials.enc or msal-cache.json.
  const mgr = new AccountManager();
  const store = (mgr as any).store;
  store.save = async () => {
    process.stderr.write('jev-attach: token refreshed in memory only; credentials.enc not written\n');
  };
  const creds = await store.get(account);
  if (!creds) {
    process.stderr.write(`jev-attach: unknown account ${account}\n`);
    process.exit(1);
  }
  if (creds.provider === 'outlook' && creds.oauth?.expiry) {
    const expiry = new Date(creds.oauth.expiry).getTime();
    if (isNaN(expiry) || expiry <= Date.now() + 5 * 60 * 1000) {
      process.stderr.write('jev-attach: SKIPPED outlook token expired; not refreshing (would write msal-cache.json)\n');
      process.stdout.write('[]');
      return;
    }
  }

  const provider = await mgr.getProvider(account);
  const email = await provider.getEmail(message);
  const safeMsg = message.replace(/[^A-Za-z0-9_-]/g, '_').slice(0, 120);
  const dir = path.resolve(QUOTES, safeMsg);
  const out: Array<Record<string, unknown>> = [];

  for (const a of email.attachments) {
    if (only && a.filename !== only) continue;
    const why = badName(a.filename);
    if (why) { out.push({ name: a.filename, skipped: why }); continue; }
    if (a.size > MAX_BYTES) { out.push({ name: a.filename, skipped: 'over 100 MB' }); continue; }
    const { data } = await provider.getAttachment(message, a.id);
    if (data.length > MAX_BYTES) { out.push({ name: a.filename, skipped: 'over 100 MB' }); continue; }
    mkdirSync(dir, { recursive: true });
    const target = path.resolve(freePath(dir, a.filename));
    if (!target.startsWith(path.resolve(QUOTES) + path.sep)) {
      out.push({ name: a.filename, skipped: 'path escaped the quotes folder' });
      continue;
    }
    writeFileSync(target, data, { flag: 'wx' });  // wx: fail rather than overwrite
    const sha256 = createHash('sha256').update(data).digest('hex');
    appendFileSync(LOG, JSON.stringify({ at: new Date().toISOString(), account, message, filename: a.filename,
      saved: target, bytes: data.length, sha256 }) + '\n');
    out.push({ name: a.filename, saved: target, bytes: data.length, sha256 });
  }
  process.stdout.write(JSON.stringify(out));
}

main().catch((err) => {
  process.stderr.write(`jev-attach: ${err?.stack ?? String(err)}\n`);
  process.exit(1);
});
