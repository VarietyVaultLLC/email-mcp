import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import type { AccountCredentials } from '../models/types.js';
import { atomicWriteFileSync, sleepSync, withLock } from './atomic-file.js';

const ALGORITHM = 'aes-256-gcm';
const KEY_LENGTH = 32;
const IV_LENGTH = 16;
const SALT_LENGTH = 32;
const PBKDF2_ITERATIONS = 100_000;

function getMachineSeed(): string {
  return `email-mcp:${os.hostname()}:${os.userInfo().username}`;
}

function deriveKey(salt: Buffer): Buffer {
  return crypto.pbkdf2Sync(getMachineSeed(), salt, PBKDF2_ITERATIONS, KEY_LENGTH, 'sha512');
}

function encrypt(plaintext: string): string {
  const salt = crypto.randomBytes(SALT_LENGTH);
  const key = deriveKey(salt);
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);

  let encrypted = cipher.update(plaintext, 'utf-8', 'hex');
  encrypted += cipher.final('hex');
  const authTag = cipher.getAuthTag();

  return JSON.stringify({
    salt: salt.toString('hex'),
    iv: iv.toString('hex'),
    authTag: authTag.toString('hex'),
    data: encrypted,
  });
}

function decrypt(ciphertext: string): string {
  const { salt, iv, authTag, data } = JSON.parse(ciphertext);
  const key = deriveKey(Buffer.from(salt, 'hex'));
  const decipher = crypto.createDecipheriv(ALGORITHM, key, Buffer.from(iv, 'hex'));
  decipher.setAuthTag(Buffer.from(authTag, 'hex'));

  let decrypted = decipher.update(data, 'hex', 'utf-8');
  decrypted += decipher.final('utf-8');
  return decrypted;
}

interface StoredData {
  accounts: Record<string, AccountCredentials>;
}

/** Stale threshold for the credentials.enc lock (see atomic-file.ts withLock). */
const CREDENTIALS_LOCK_STALE_MS = 10_000;
const READ_RETRY_ATTEMPTS = 5;
const READ_RETRY_DELAY_MS = 50;

export class CredentialStore {
  private filePath: string;

  constructor(dir?: string) {
    const baseDir = dir ?? path.join(os.homedir(), '.email-mcp');
    if (!fs.existsSync(baseDir)) {
      fs.mkdirSync(baseDir, { recursive: true, mode: 0o700 });
    }
    this.filePath = path.join(baseDir, 'credentials.enc');
  }

  private read(): StoredData {
    if (!fs.existsSync(this.filePath)) {
      return { accounts: {} };
    }
    let lastErr: unknown;
    for (let attempt = 1; attempt <= READ_RETRY_ATTEMPTS; attempt++) {
      try {
        const raw = fs.readFileSync(this.filePath, 'utf-8');
        if (raw.length === 0) {
          throw new Error('credentials file is empty');
        }
        const json = decrypt(raw);
        return JSON.parse(json);
      } catch (err) {
        lastErr = err;
        if (attempt < READ_RETRY_ATTEMPTS) {
          sleepSync(READ_RETRY_DELAY_MS);
        }
      }
    }
    throw lastErr;
  }

  private write(data: StoredData): void {
    const json = JSON.stringify(data);
    const encrypted = encrypt(json);
    atomicWriteFileSync(this.filePath, encrypted, 0o600);
  }

  async save(creds: AccountCredentials): Promise<void> {
    withLock(this.filePath, CREDENTIALS_LOCK_STALE_MS, () => {
      const data = this.read();
      data.accounts[creds.id] = creds;
      this.write(data);
    });
  }

  async get(id: string): Promise<AccountCredentials | null> {
    const data = this.read();
    return data.accounts[id] ?? null;
  }

  async list(): Promise<AccountCredentials[]> {
    const data = this.read();
    return Object.values(data.accounts);
  }

  async remove(id: string): Promise<void> {
    withLock(this.filePath, CREDENTIALS_LOCK_STALE_MS, () => {
      const data = this.read();
      delete data.accounts[id];
      this.write(data);
    });
  }
}
