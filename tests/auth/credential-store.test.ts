import { describe, it, expect, beforeAll, afterAll, beforeEach, afterEach } from 'vitest';
import { CredentialStore } from '../../src/auth/credential-store.js';
import { ProviderType } from '../../src/models/types.js';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { build } from 'esbuild';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const WORKER_SRC_PATH = path.join(__dirname, 'concurrent-save-worker.ts');

// The worker imports credential-store.js (bundler-style resolution to the
// .ts source, same convention as the rest of this codebase), which node
// cannot resolve directly. Bundle it once with esbuild -- the project's own
// build tool -- into a single-file script, then spawn that as a real,
// independent OS process per save, the same way the live MCP server
// processes and CLI tools (watcher-fetch, inbound-scan, jev-fetch) share
// credentials.enc.
let workerBundlePath: string;

beforeAll(async () => {
  workerBundlePath = path.join(
    os.tmpdir(),
    `email-mcp-concurrent-save-worker-${process.pid}.mjs`,
  );
  await build({
    entryPoints: [WORKER_SRC_PATH],
    outfile: workerBundlePath,
    bundle: true,
    platform: 'node',
    format: 'esm',
    target: 'node20',
  });
});

afterAll(() => {
  if (workerBundlePath) {
    fs.rmSync(workerBundlePath, { force: true });
  }
});

/** Spawns the built worker as a real child process (not a vitest worker thread). */
function spawnSaveWorker(dir: string, accountId: string): Promise<{ code: number | null; stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [workerBundlePath, dir, accountId], {
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => (stdout += chunk.toString()));
    child.stderr.on('data', (chunk) => (stderr += chunk.toString()));
    child.on('error', reject);
    child.on('close', (code) => resolve({ code, stdout, stderr }));
  });
}

describe('CredentialStore', () => {
  let store: CredentialStore;
  let testDir: string;

  beforeEach(() => {
    testDir = fs.mkdtempSync(path.join(os.tmpdir(), 'email-mcp-test-'));
    store = new CredentialStore(testDir);
  });

  afterEach(() => {
    fs.rmSync(testDir, { recursive: true, force: true });
  });

  it('saves and loads an account', async () => {
    const creds = {
      id: 'test-1',
      name: 'Test Gmail',
      provider: ProviderType.Gmail as const,
      email: 'test@gmail.com',
      oauth: {
        access_token: 'at-123',
        refresh_token: 'rt-456',
        expiry: '2026-12-31T00:00:00Z',
      },
    };

    await store.save(creds);
    const loaded = await store.get('test-1');
    expect(loaded).toEqual(creds);
  });

  it('lists all accounts', async () => {
    await store.save({
      id: 'a1',
      name: 'Gmail',
      provider: ProviderType.Gmail as const,
      email: 'a@gmail.com',
      oauth: { access_token: 'x', refresh_token: 'y', expiry: '' },
    });
    await store.save({
      id: 'a2',
      name: 'iCloud',
      provider: ProviderType.ICloud as const,
      email: 'b@icloud.com',
      password: { password: 'p', host: 'imap.mail.me.com', port: 993, tls: true },
    });

    const accounts = await store.list();
    expect(accounts).toHaveLength(2);
    expect(accounts.map((a) => a.id).sort()).toEqual(['a1', 'a2']);
  });

  it('removes an account', async () => {
    await store.save({
      id: 'del-1',
      name: 'ToDelete',
      provider: ProviderType.IMAP as const,
      email: 'x@test.com',
      password: { password: 'p', host: 'imap.test.com', port: 993, tls: true },
    });

    await store.remove('del-1');
    const loaded = await store.get('del-1');
    expect(loaded).toBeNull();
  });

  it('returns null for nonexistent account', async () => {
    const loaded = await store.get('nope');
    expect(loaded).toBeNull();
  });

  it('persists across instances', async () => {
    await store.save({
      id: 'persist-1',
      name: 'Persist',
      provider: ProviderType.Gmail as const,
      email: 'p@gmail.com',
      oauth: { access_token: 'a', refresh_token: 'r', expiry: '' },
    });

    const store2 = new CredentialStore(testDir);
    const loaded = await store2.get('persist-1');
    expect(loaded?.email).toBe('p@gmail.com');
  });

  it('encrypts the file on disk', async () => {
    await store.save({
      id: 'enc-1',
      name: 'Encrypted',
      provider: ProviderType.Gmail as const,
      email: 'enc@gmail.com',
      oauth: { access_token: 'secret-token', refresh_token: 'secret-refresh', expiry: '' },
    });

    const filePath = path.join(testDir, 'credentials.enc');
    const raw = fs.readFileSync(filePath, 'utf-8');
    expect(raw).not.toContain('secret-token');
    expect(raw).not.toContain('secret-refresh');
  });

  it('updates an existing account', async () => {
    const creds = {
      id: 'upd-1',
      name: 'Original',
      provider: ProviderType.Gmail as const,
      email: 'u@gmail.com',
      oauth: { access_token: 'old', refresh_token: 'r', expiry: '' },
    };
    await store.save(creds);
    await store.save({ ...creds, name: 'Updated', oauth: { access_token: 'new', refresh_token: 'r', expiry: '' } });

    const loaded = await store.get('upd-1');
    expect(loaded?.name).toBe('Updated');
    expect(loaded?.oauth?.access_token).toBe('new');

    const all = await store.list();
    expect(all).toHaveLength(1);
  });

  it('leaves no temp files behind after a save', async () => {
    await store.save({
      id: 'tmp-check',
      name: 'Tmp Check',
      provider: ProviderType.Gmail as const,
      email: 'tmp@gmail.com',
      oauth: { access_token: 'a', refresh_token: 'r', expiry: '' },
    });

    const leftovers = fs.readdirSync(testDir).filter((f) => f.endsWith('.tmp'));
    expect(leftovers).toEqual([]);
  });

  it('releases the lock file after a save completes', async () => {
    await store.save({
      id: 'lock-check',
      name: 'Lock Check',
      provider: ProviderType.Gmail as const,
      email: 'lock@gmail.com',
      oauth: { access_token: 'a', refresh_token: 'r', expiry: '' },
    });

    const lockPath = path.join(testDir, 'credentials.enc.lock');
    expect(fs.existsSync(lockPath)).toBe(false);
  });

  it('releases the lock after remove() too', async () => {
    await store.save({
      id: 'rm-lock-check',
      name: 'Rm Lock Check',
      provider: ProviderType.Gmail as const,
      email: 'rmlock@gmail.com',
      oauth: { access_token: 'a', refresh_token: 'r', expiry: '' },
    });
    await store.remove('rm-lock-check');

    const lockPath = path.join(testDir, 'credentials.enc.lock');
    expect(fs.existsSync(lockPath)).toBe(false);
  });

  it('breaks a stale lock instead of wedging forever', async () => {
    const lockPath = path.join(testDir, 'credentials.enc.lock');
    fs.writeFileSync(lockPath, '');
    // Back-date the lock well past the 10s staleness threshold.
    const old = new Date(Date.now() - 60_000);
    fs.utimesSync(lockPath, old, old);

    await store.save({
      id: 'stale-lock',
      name: 'Stale Lock',
      provider: ProviderType.Gmail as const,
      email: 'stale@gmail.com',
      oauth: { access_token: 'a', refresh_token: 'r', expiry: '' },
    });

    const loaded = await store.get('stale-lock');
    expect(loaded?.email).toBe('stale@gmail.com');
    expect(fs.existsSync(lockPath)).toBe(false);
  });

  it('leaves a decryptable file after a save even under a pre-existing lock contender', async () => {
    // Simulate a short-lived lock held by another process: create it, then
    // release it shortly after from a separate OS process (save() blocks
    // this thread synchronously while it waits, so an in-process timer
    // would never get a chance to fire), and make sure our save still lands.
    const lockPath = path.join(testDir, 'credentials.enc.lock');
    const fd = fs.openSync(lockPath, 'wx');
    fs.closeSync(fd);
    const releaser = spawn(
      process.execPath,
      ['-e', `setTimeout(()=>{try{require('fs').rmSync(process.argv[1],{force:true})}catch{}},100)`, lockPath],
      { stdio: 'ignore' },
    );

    await store.save({
      id: 'contended',
      name: 'Contended',
      provider: ProviderType.Gmail as const,
      email: 'contended@gmail.com',
      oauth: { access_token: 'a', refresh_token: 'r', expiry: '' },
    });

    const loaded = await store.get('contended');
    expect(loaded?.email).toBe('contended@gmail.com');
  });

  it('survives concurrent saves from separate OS processes without losing updates', async () => {
    const accountIds = ['proc-a', 'proc-b', 'proc-c', 'proc-d', 'proc-e'];

    const results = await Promise.all(
      accountIds.map((id) => spawnSaveWorker(testDir, id)),
    );

    for (const result of results) {
      expect(result.code, `worker stderr: ${result.stderr}`).toBe(0);
    }

    // Fresh store instance so we're reading what actually landed on disk,
    // not anything cached in-process.
    const verifyStore = new CredentialStore(testDir);
    const all = await verifyStore.list();
    const ids = all.map((a) => a.id).sort();
    expect(ids).toEqual([...accountIds].sort());

    // The file must decrypt cleanly (no truncated/partial write survived).
    const filePath = path.join(testDir, 'credentials.enc');
    const raw = fs.readFileSync(filePath, 'utf-8');
    expect(() => JSON.parse(raw)).not.toThrow();

    const leftovers = fs.readdirSync(testDir).filter((f) => f.endsWith('.tmp') || f.endsWith('.lock'));
    expect(leftovers).toEqual([]);
  }, 30_000);
});
