import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

/** Synchronous sleep used only for short, bounded retry backoff. */
export function sleepSync(ms: number): void {
  const buf = new SharedArrayBuffer(4);
  const view = new Int32Array(buf);
  Atomics.wait(view, 0, 0, ms);
}

/**
 * Writes `contents` to `filePath` atomically: write to a temp file in the
 * same directory, then rename over the target. Retries the rename on
 * Windows' transient EPERM/EBUSY/EACCES (another process has the file open)
 * instead of failing outright. Shared by any writer of a file under
 * ~/.email-mcp/ that other MCP server processes or CLI tools may read
 * concurrently (credentials.enc, msal-cache.json).
 */
export function atomicWriteFileSync(
  filePath: string,
  contents: string,
  mode: number = 0o600,
): void {
  const dir = path.dirname(filePath);
  const tmpPath = path.join(
    dir,
    `${path.basename(filePath)}.${process.pid}.${crypto.randomBytes(4).toString('hex')}.tmp`,
  );

  fs.writeFileSync(tmpPath, contents, { mode });

  const maxAttempts = 10;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      fs.renameSync(tmpPath, filePath);
      return;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      const retryable = code === 'EPERM' || code === 'EBUSY' || code === 'EACCES';
      if (!retryable || attempt === maxAttempts) {
        try {
          fs.unlinkSync(tmpPath);
        } catch {
          // best effort cleanup
        }
        throw err;
      }
      sleepSync(Math.min(20 * attempt, 200));
    }
  }
}

const LOCK_RETRY_TOTAL_MS = 10_000;
const LOCK_RETRY_DELAY_MS = 25;

/**
 * Acquires an exclusive cross-process lock file at `lockPath` (created via
 * `wx` so only one process can succeed). Retries on EEXIST for up to
 * LOCK_RETRY_TOTAL_MS. A lock whose mtime is older than `staleMs` is treated
 * as abandoned (a crashed process, or a bug that acquired without releasing)
 * and is removed so callers don't wedge forever — this applies equally to a
 * lock left behind by the *same* process, since the check is purely
 * file-mtime based, not PID based.
 */
export function acquireLockSync(lockPath: string, staleMs: number): void {
  const deadline = Date.now() + LOCK_RETRY_TOTAL_MS;

  for (;;) {
    try {
      const fd = fs.openSync(lockPath, 'wx');
      fs.closeSync(fd);
      return;
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== 'EEXIST') {
        throw err;
      }
      try {
        const stat = fs.statSync(lockPath);
        if (Date.now() - stat.mtimeMs > staleMs) {
          fs.rmSync(lockPath, { force: true });
          continue;
        }
      } catch {
        // lock disappeared between openSync and statSync; retry loop below
      }
      if (Date.now() >= deadline) {
        throw new Error(`Timed out waiting for lock: ${lockPath}`);
      }
      sleepSync(LOCK_RETRY_DELAY_MS);
    }
  }
}

/** Releases a lock acquired via {@link acquireLockSync}. Best effort. */
export function releaseLockSync(lockPath: string): void {
  try {
    fs.rmSync(lockPath, { force: true });
  } catch {
    // best effort release
  }
}

/**
 * Acquires the lock at `${filePath}.lock`, runs `fn`, and always releases
 * the lock (even if `fn` throws). Convenience wrapper for the common case
 * where the whole critical section is a single synchronous call — callers
 * whose critical section spans multiple async callbacks (e.g. MSAL's
 * beforeCacheAccess/afterCacheAccess pair) should call
 * {@link acquireLockSync}/{@link releaseLockSync} directly instead.
 */
export function withLock<T>(filePath: string, staleMs: number, fn: () => T): T {
  const lockPath = `${filePath}.lock`;
  acquireLockSync(lockPath, staleMs);
  try {
    return fn();
  } finally {
    releaseLockSync(lockPath);
  }
}
