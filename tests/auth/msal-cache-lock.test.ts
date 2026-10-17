import { describe, it, expect, beforeAll, afterAll, beforeEach, afterEach } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { build } from 'esbuild';
import { createEncryptedFileCachePlugin } from '../../src/providers/outlook/auth.js';
import { decryptWithMachineKey } from '../../src/auth/credential-store.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const WORKER_SRC_PATH = path.join(__dirname, 'msal-cache-lock-worker.ts');

// The worker pulls in @azure/msal-node, whose CJS dependencies esbuild cannot
// fully bundle into ESM, so msal-node stays external (as in build.mjs) and the
// bundle is written inside the repo where node_modules resolution finds it.
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
  if (workerBundlePath) fs.rmSync(workerBundlePath, { force: true });
});

function spawnCacheWorker(cacheFilePath: string, workerId: string, cycles: number): Promise<{ code: number | null; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [workerBundlePath, cacheFilePath, workerId, String(cycles)], {
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stderr = '';
    child.stderr.on('data', (chunk) => (stderr += chunk.toString()));
    child.on('error', reject);
    child.on('close', (code) => resolve({ code, stderr }));
  });
}

/** Minimal fake TokenCacheContext: the plugin only touches these members. */
function makeFakeContext() {
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
  return { context: context as any, store };
}

const readCache = (file: string) => JSON.parse(decryptWithMachineKey(fs.readFileSync(file, 'utf-8')).plaintext);

describe('MSAL cache plugin locking', () => {
  let testDir: string;
  let cacheFilePath: string;

  beforeEach(() => {
    testDir = fs.mkdtempSync(path.join(os.tmpdir(), 'email-mcp-msal-lock-'));
    cacheFilePath = path.join(testDir, 'msal-cache.enc');
  });

  afterEach(() => {
    fs.rmSync(testDir, { recursive: true, force: true });
  });

  it('writes the cache and releases the lock after a normal cycle', async () => {
    const plugin = createEncryptedFileCachePlugin(cacheFilePath);
    const { context, store } = makeFakeContext();

    await plugin.beforeCacheAccess(context);
    store.data.foo = 'bar';
    context.cacheHasChanged = true;
    await plugin.afterCacheAccess(context);

    expect(fs.existsSync(`${cacheFilePath}.lock`)).toBe(false);
    expect(readCache(cacheFilePath)).toEqual({ foo: 'bar' });
  });

  it('releases the lock without writing when the cache did not change', async () => {
    const plugin = createEncryptedFileCachePlugin(cacheFilePath);
    const { context } = makeFakeContext();

    await plugin.beforeCacheAccess(context);
    await plugin.afterCacheAccess(context);

    expect(fs.existsSync(cacheFilePath)).toBe(false);
    expect(fs.existsSync(`${cacheFilePath}.lock`)).toBe(false);
  });

  it('releases the lock when the cache cannot be read', async () => {
    fs.writeFileSync(cacheFilePath, 'not an encrypted cache');
    const plugin = createEncryptedFileCachePlugin(cacheFilePath);
    const { context } = makeFakeContext();

    await expect(plugin.beforeCacheAccess(context)).rejects.toThrow();
    // MSAL still calls afterCacheAccess after a failed before; it must not
    // write, and must not release a lock it no longer holds.
    await plugin.afterCacheAccess(context);
    expect(fs.existsSync(`${cacheFilePath}.lock`)).toBe(false);
    expect(fs.readFileSync(cacheFilePath, 'utf-8')).toBe('not an encrypted cache');
  });

  it('breaks a stale lock instead of wedging forever', async () => {
    const lockPath = `${cacheFilePath}.lock`;
    fs.writeFileSync(lockPath, '');
    const old = new Date(Date.now() - 45_000);
    fs.utimesSync(lockPath, old, old);

    const plugin = createEncryptedFileCachePlugin(cacheFilePath);
    const { context, store } = makeFakeContext();
    await plugin.beforeCacheAccess(context);
    store.data.unstuck = true;
    context.cacheHasChanged = true;
    await plugin.afterCacheAccess(context);

    expect(fs.existsSync(lockPath)).toBe(false);
    expect(readCache(cacheFilePath)).toEqual({ unstuck: true });
  });

  it('keeps every key when separate OS processes cycle the cache at once', async () => {
    const workerIds = ['w1', 'w2', 'w3', 'w4'];
    const cyclesEach = 3;
    const results = await Promise.all(workerIds.map((id) => spawnCacheWorker(cacheFilePath, id, cyclesEach)));
    for (const result of results) {
      expect(result.code, `worker stderr: ${result.stderr}`).toBe(0);
    }

    const expected = workerIds.flatMap((id) => Array.from({ length: cyclesEach }, (_, i) => `${id}-${i}`));
    expect(Object.keys(readCache(cacheFilePath)).sort()).toEqual(expected.sort());
    const leftovers = fs.readdirSync(testDir).filter((f) => f.endsWith('.tmp') || f.endsWith('.lock'));
    expect(leftovers).toEqual([]);
  }, 30_000);
});
