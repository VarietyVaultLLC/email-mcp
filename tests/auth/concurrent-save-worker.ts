/**
 * Standalone worker used by credential-store.test.ts to exercise real
 * multi-process concurrency against a shared CredentialStore directory.
 * The test bundles this file with esbuild (the project's own build tool)
 * into a single-file script and spawns it with `node` (not through vitest)
 * so each invocation is a genuinely separate OS process, matching how the
 * MCP server processes and CLI tools (watcher-fetch, inbound-scan,
 * jev-fetch) share credentials.enc.
 *
 * Usage: node <bundled-worker>.mjs <dir> <accountId>
 */
import { CredentialStore } from '../../src/auth/credential-store.js';

async function main(): Promise<void> {
  const [dir, accountId] = process.argv.slice(2);
  if (!dir || !accountId) {
    throw new Error('usage: concurrent-save-worker.ts <dir> <accountId>');
  }

  const store = new CredentialStore(dir);
  await store.save({
    id: accountId,
    name: `Worker ${accountId}`,
    provider: 'gmail',
    email: `${accountId}@example.com`,
    oauth: {
      access_token: `at-${accountId}`,
      refresh_token: `rt-${accountId}`,
      expiry: '',
    },
  });
  process.stdout.write(`OK ${accountId}\n`);
}

main().catch((err) => {
  process.stderr.write(`WORKER_ERROR ${String(err)}\n`);
  process.exit(1);
});
