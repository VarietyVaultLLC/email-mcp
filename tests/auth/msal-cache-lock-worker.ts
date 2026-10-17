/**
 * Standalone worker used by msal-cache-lock.test.ts to exercise the MSAL
 * cache plugin's locking under real multi-process concurrency. Bundled with
 * esbuild and spawned as an independent `node` process, the way several
 * server processes and CLI tools race to refresh Outlook tokens.
 *
 * Runs `cycles` before -> mutate -> after cycles against one shared cache
 * file, adding a unique key per cycle, through a minimal fake
 * TokenCacheContext (the plugin only touches deserialize, serialize and
 * cacheHasChanged).
 *
 * Usage: node <bundled-worker>.mjs <cacheFilePath> <workerId> <cycles>
 */
import { createEncryptedFileCachePlugin } from '../../src/providers/outlook/auth.js';
import { sleepSync } from '../../src/auth/atomic-file.js';

async function main(): Promise<void> {
  const [cacheFilePath, workerId, cyclesArg] = process.argv.slice(2);
  if (!cacheFilePath || !workerId) {
    throw new Error('usage: msal-cache-lock-worker.ts <cacheFilePath> <workerId> <cycles>');
  }
  const cycles = parseInt(cyclesArg ?? '1', 10);
  const plugin = createEncryptedFileCachePlugin(cacheFilePath);

  for (let i = 0; i < cycles; i++) {
    const store = { data: {} as Record<string, unknown> };
    const context = {
      cacheHasChanged: false,
      tokenCache: {
        deserialize(data: string) {
          store.data = data.length > 0 ? JSON.parse(data) : {};
        },
        serialize() {
          return JSON.stringify(store.data);
        },
      },
    };

    await plugin.beforeCacheAccess(context as any);
    // A small random delay between read and write widens the race window.
    sleepSync(Math.floor(Math.random() * 15));
    store.data[`${workerId}-${i}`] = { pid: process.pid, i };
    context.cacheHasChanged = true;
    await plugin.afterCacheAccess(context as any);
  }

  process.stdout.write(`OK ${workerId}\n`);
}

main().catch((err) => {
  process.stderr.write(`WORKER_ERROR ${String(err)}\n`);
  process.exit(1);
});
