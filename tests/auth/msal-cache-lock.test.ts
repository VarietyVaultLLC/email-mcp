import { describe, it, expect, beforeAll, afterAll, beforeEach, afterEach } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { build } from 'esbuild';
import { createFileCachePlugin } from '../../src/providers/outlook/auth.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const WORKER_SRC_PATH = path.join(__dirname, 'msal-cache-lock-worker.ts');

// Bundle the worker with esbuild (the project's own build tool), the same
// style used by the credential-store concurrency test, then spawn it as a
// real, independent `node` child process. Unlike that worker, this one pulls
// in @azure/msal-node, which has CJS dependencies (jsonwebtoken/jws) that
// esbuild cannot fully bundle into an ESM output (dynamic `require()` of a
// node builtin). So, matching how build.mjs itself treats it, msal-node is
// left external here rather than bundled -- which means the bundle must live
// somewhere node_modules resolution can still find it, hence writing it
// inside the repo (tests/auth/) rather than os.tmpdir().
let workerBundlePath: string;

beforeAll(async () => {
  workerBundlePath = path.join(__dirname, `.msal-cache-lock-worker.${process.pid}.worker-bundle.mjs`);
  await build({
    entryPoints: [WORKER_SRC_PATH],
    outfile: workerBundlePath,
    bundle: true,
    platform: 'node',
    format: 'esm',
    target: 'node20',
    external: ['@azure/msal-node'],
  });
}, 30_000);

afterAll(() => {
  if (workerBundlePath) {
    fs.rmSync(workerBundlePath, { force: true });
  }
});

function spawnCacheWorker(
  cacheFilePath: string,
  workerId: string,
  cycles: number,
): Promise<{ code: number | null; stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [workerBundlePath, cacheFilePath, workerId, String(cycles)],
      { stdio: ['ignore', 'pipe', 'pipe'] },
    );
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => (stdout += chunk.toString()));
    child.stderr.on('data', (chunk) => (stderr += chunk.toString()));
    child.on('error', reject);
    child.on('close', (code) => resolve({ code, stdout, stderr }));
  });
}

/** Minimal fake TokenCacheContext -- the plugin only touches these members. */
function makeFakeContext(initial: Record<string, unknown> = {}) {
  const store = { data: { ...initial } };
  return {
    context: {
      cacheHasChanged: false,
      tokenCache: {
        deserialize(data: string) {
          store.data = data.length > 0 ? JSON.parse(data) : {};
        },
        serialize() {
          return JSON.stringify(store.data);
        },
      },
    },
    store,
  };
}

describe('MSAL file cache plugin locking', () => {
  let testDir: string;
  let cacheFilePath: string;

  beforeEach(() => {
    testDir = fs.mkdtempSync(path.join(os.tmpdir(), 'email-mcp-msal-test-'));
    cacheFilePath = path.join(testDir, 'msal-cache.json');
  });

  afterEach(() => {
    fs.rmSync(testDir, { recursive: true, force: true });
  });

  it('writes the cache and releases the lock after a normal cycle', async () => {
    const plugin = createFileCachePlugin(cacheFilePath);
    const { context, store } = makeFakeContext();

    await (plugin.beforeCacheAccess as any)(context);
    store.data.foo = 'bar';
    context.cacheHasChanged = true;
    await (plugin.afterCacheAccess as any)(context);

    expect(fs.existsSync(`${cacheFilePath}.lock`)).toBe(false);
    const raw = fs.readFileSync(cacheFilePath, 'utf-8');
    expect(JSON.parse(raw)).toEqual({ foo: 'bar' });
  });

  it('does not write the file when cacheHasChanged is false', async () => {
    const plugin = createFileCachePlugin(cacheFilePath);
    const { context } = makeFakeContext();

    await (plugin.beforeCacheAccess as any)(context);
    // context.cacheHasChanged left false
    await (plugin.afterCacheAccess as any)(context);

    expect(fs.existsSync(cacheFilePath)).toBe(false);
    expect(fs.existsSync(`${cacheFilePath}.lock`)).toBe(false);
  });

  it('releases the lock if beforeCacheAccess throws after acquiring it', async () => {
    // Write a file that exists() sees but readFileSync/JSON.parse chokes on
    // by making deserialize throw, simulating a corrupt read.
    fs.writeFileSync(cacheFilePath, '{"ok":true}');
    const plugin = createFileCachePlugin(cacheFilePath);
    const context = {
      cacheHasChanged: false,
      tokenCache: {
        deserialize() {
          throw new Error('boom: simulated corrupt cache');
        },
        serialize() {
          return '{}';
        },
      },
    };

    await expect((plugin.beforeCacheAccess as any)(context)).rejects.toThrow(
      'boom: simulated corrupt cache',
    );
    expect(fs.existsSync(`${cacheFilePath}.lock`)).toBe(false);
  });

  it('breaks a stale MSAL lock instead of wedging forever', async () => {
    const lockPath = `${cacheFilePath}.lock`;
    fs.writeFileSync(lockPath, '');
    // Back-date well past the 30s MSAL staleness threshold.
    const old = new Date(Date.now() - 45_000);
    fs.utimesSync(lockPath, old, old);

    const plugin = createFileCachePlugin(cacheFilePath);
    const { context, store } = makeFakeContext();

    await (plugin.beforeCacheAccess as any)(context);
    store.data.unstuck = true;
    context.cacheHasChanged = true;
    await (plugin.afterCacheAccess as any)(context);

    expect(fs.existsSync(lockPath)).toBe(false);
    const raw = fs.readFileSync(cacheFilePath, 'utf-8');
    expect(JSON.parse(raw)).toEqual({ unstuck: true });
  });

  it(
    'survives concurrent before/mutate/after cycles from separate OS processes without losing keys',
    async () => {
      const workerIds = ['w1', 'w2', 'w3', 'w4'];
      const cyclesEach = 3;

      const results = await Promise.all(
        workerIds.map((id) => spawnCacheWorker(cacheFilePath, id, cyclesEach)),
      );

      for (const result of results) {
        expect(result.code, `worker stderr: ${result.stderr}`).toBe(0);
      }

      const raw = fs.readFileSync(cacheFilePath, 'utf-8');
      const parsed = JSON.parse(raw);

      const expectedKeys = workerIds.flatMap((id) =>
        Array.from({ length: cyclesEach }, (_, i) => `${id}-${i}`),
      );
      expect(Object.keys(parsed).sort()).toEqual(expectedKeys.sort());

      const leftovers = fs
        .readdirSync(testDir)
        .filter((f) => f.endsWith('.tmp') || f.endsWith('.lock'));
      expect(leftovers).toEqual([]);
    },
    30_000,
  );
});
