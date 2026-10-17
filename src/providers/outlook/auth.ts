import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import {
  PublicClientApplication,
  CryptoProvider,
  type AuthenticationResult,
  type ICachePlugin,
  type TokenCacheContext,
} from '@azure/msal-node';
import { atomicWriteFileSync, acquireLockSync, releaseLockSync } from '../../auth/atomic-file.js';

export interface OutlookAuthResult {
  accessToken: string;
  expiresOn: Date | null;
  idToken?: string;
  account?: unknown;
  /** MSAL home account ID for subsequent silent token acquisition */
  homeAccountId?: string;
}

/** Default directory for MSAL cache file */
const DEFAULT_CACHE_DIR = path.join(os.homedir(), '.email-mcp');
const DEFAULT_CACHE_FILE = path.join(DEFAULT_CACHE_DIR, 'msal-cache.json');

/**
 * Stale threshold for the msal-cache.json lock. Longer than the credentials
 * lock (10s) because a cache access cycle here can involve a network round
 * trip to Microsoft (silent token refresh) between beforeCacheAccess and
 * afterCacheAccess, not just a local read-modify-write.
 */
const MSAL_LOCK_STALE_MS = 30_000;

/**
 * Creates a file-based ICachePlugin for persisting MSAL's token cache.
 * This ensures refresh tokens survive process restarts.
 *
 * beforeCacheAccess/afterCacheAccess bracket one logical cache access: the
 * lock is acquired in beforeCacheAccess (before reading) and released in
 * afterCacheAccess (after writing, if the cache changed), so no other
 * process — MCP server instance or CLI tool — can read or write
 * msal-cache.json in between. If beforeCacheAccess itself throws after
 * acquiring the lock (e.g. a corrupt/partial read), it releases before
 * rethrowing. afterCacheAccess always releases via try/finally, whether or
 * not it wrote. A lock left held (crash, or a bug that skips the release)
 * is broken once it's older than MSAL_LOCK_STALE_MS, including a lock left
 * by this same process — the check is purely file-mtime based, so a later
 * acquire can never deadlock forever.
 */
export function createFileCachePlugin(cacheFilePath: string): ICachePlugin {
  const lockPath = `${cacheFilePath}.lock`;

  return {
    async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
      acquireLockSync(lockPath, MSAL_LOCK_STALE_MS);
      try {
        if (fs.existsSync(cacheFilePath)) {
          const data = fs.readFileSync(cacheFilePath, 'utf-8');
          context.tokenCache.deserialize(data);
        }
      } catch (err) {
        releaseLockSync(lockPath);
        throw err;
      }
    },
    async afterCacheAccess(context: TokenCacheContext): Promise<void> {
      try {
        if (context.cacheHasChanged) {
          const dir = path.dirname(cacheFilePath);
          if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
          }
          atomicWriteFileSync(cacheFilePath, context.tokenCache.serialize(), 0o600);
        }
      } finally {
        releaseLockSync(lockPath);
      }
    },
  };
}

export class OutlookAuth {
  static readonly AUTHORITY = 'https://login.microsoftonline.com/common';
  static readonly SCOPES = ['Mail.ReadWrite', 'Mail.Send', 'offline_access'];
  /**
   * Requested separately (incremental consent) so mail keeps working on
   * tokens consented before calendar support existed. Granted once via
   * `node dist/cli/grant-calendar.js`.
   */
  static readonly CALENDAR_SCOPES = ['Calendars.ReadWrite'];

  private pca: InstanceType<typeof PublicClientApplication>;
  private cryptoProvider: InstanceType<typeof CryptoProvider>;

  constructor(clientId: string, cacheFilePath?: string) {
    this.pca = new PublicClientApplication({
      auth: {
        clientId,
        authority: OutlookAuth.AUTHORITY,
      },
      cache: {
        cachePlugin: createFileCachePlugin(cacheFilePath ?? DEFAULT_CACHE_FILE),
      },
    });
    this.cryptoProvider = new CryptoProvider();
  }

  async getAuthUrl(
    redirectUri: string,
    scopes: string[] = OutlookAuth.SCOPES
  ): Promise<{ url: string; codeVerifier: string }> {
    const { verifier, challenge } = await this.cryptoProvider.generatePkceCodes();

    let url = await this.pca.getAuthCodeUrl({
      scopes,
      redirectUri,
      codeChallenge: challenge,
      codeChallengeMethod: 'S256',
    });

    // Force account picker — MSAL strips the prompt param, so append manually
    url += '&prompt=select_account';

    return { url, codeVerifier: verifier };
  }

  async exchangeCode(
    code: string,
    codeVerifier: string,
    redirectUri: string,
    scopes: string[] = OutlookAuth.SCOPES
  ): Promise<OutlookAuthResult> {
    const result: AuthenticationResult = await this.pca.acquireTokenByCode({
      code,
      codeVerifier,
      scopes,
      redirectUri,
    });

    return {
      accessToken: result.accessToken,
      expiresOn: result.expiresOn,
      idToken: result.idToken,
      account: result.account,
      homeAccountId: result.account?.homeAccountId,
    };
  }

  /**
   * Silently refresh the access token using MSAL's persisted cache.
   * This is the preferred refresh method — it uses the cached refresh token
   * automatically via acquireTokenSilent().
   */
  async refreshTokenSilent(
    homeAccountId: string,
    scopes: string[] = OutlookAuth.SCOPES
  ): Promise<OutlookAuthResult> {
    const account = await this.pca.getTokenCache().getAccountByHomeId(homeAccountId);

    if (!account) {
      throw new Error(
        'No cached MSAL account found — re-authenticate via setup wizard'
      );
    }

    const result = await this.pca.acquireTokenSilent({
      account,
      scopes,
    });

    return {
      accessToken: result.accessToken,
      expiresOn: result.expiresOn,
      homeAccountId: result.account?.homeAccountId ?? homeAccountId,
    };
  }

  /**
   * @deprecated Use refreshTokenSilent() instead.
   * Kept for backward compatibility with credentials that have a stored refresh token.
   */
  async refreshToken(refreshToken: string): Promise<OutlookAuthResult> {
    const result = await this.pca.acquireTokenByRefreshToken({
      refreshToken,
      scopes: OutlookAuth.SCOPES,
    });

    if (!result) {
      throw new Error('Failed to refresh Outlook token');
    }

    return {
      accessToken: result.accessToken,
      expiresOn: result.expiresOn,
    };
  }
}
