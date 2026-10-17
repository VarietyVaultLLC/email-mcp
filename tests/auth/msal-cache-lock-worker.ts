/**
 * Standalone worker used by msal-cache-lock.test.ts to exercise the MSAL
 * file cache plugin's locking under real multi-process concurrency. Bundled
 * with esbuild (self-contained, so it can run from any temp directory) and
 * spawned as an independent `node` child process, the same way real MCP
 * server processes race to refresh Outlook tokens and write msal-cache.json.
 *
 * Runs `cycles` before -> mutate -> after cycles against one shared cache
 * file, adding a unique key per cycle, using a minimal fake
 * TokenCacheContext (deserialize/serialize/cacheHasChanged) instead of a
 * real MSAL TokenCache -- the file cache plugin only ever touches those
 * three members.
 *
 * Usage: node <bundled-worker>.mjs <cacheFilePath> <workerId> <cycles>
 */
import { createFileCachePlugin } from '../../src/providers/outlook/auth.js';
import { sleepSync } from '../../src/auth/atomic-file.js';

interface FakeTokenCache {
  deserialize(data: string): void;
  serialize(): string;
}

interface FakeContext {
  cacheHasChanged: boolean;
  tokenCache: FakeTokenCache;
}

function makeContext(store: { data: Record<string, unknown> }): FakeContext {
  return {
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
}

async function main(): Promise<void> {
  const [cacheFilePath, workerId, cyclesArg] = process.argv.slice(2);
  if (!cacheFilePath || !workerId) {
    throw new Error('usage: msal-cache-lock-worker.ts <cacheFilePath> <workerId> <cycles>');
  }
  const cycles = parseInt(cyclesArg ?? '1', 10);

  const plugin = createFileCachePlugin(cacheFilePath);

  for (let i = 0; i < cycles; i++) {
    const store = { data: {} as Record<string, unknown> };
    const context = makeContext(store);

    // @ts-expect-error -- minimal fake stands in for MSAL's TokenCacheContext
    await plugin.beforeCacheAccess(context);

    // A tiny random delay between read and write widens the race window so
    // concurrent processes are more likely to actually overlap.
    sleepSync(Math.floor(Math.random() * 15));

    store.data[`${workerId}-${i}`] = { pid: process.pid, i, ts: Date.now() };
    context.cacheHasChanged = true;

    // @ts-expect-error -- same fake context as above
    await plugin.afterCacheAccess(context);
  }

  process.stdout.write(`OK ${workerId}\n`);
}

main().catch((err) => {
  process.stderr.write(`WORKER_ERROR ${String(err)}\n`);
  process.exit(1);
});
