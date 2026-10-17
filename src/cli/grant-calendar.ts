/**
 * One-time incremental consent for the Outlook calendar scope.
 *
 * Runs the same browser auth-code flow as the setup wizard, but asks for the
 * mail scopes PLUS Calendars.ReadWrite. MSAL stores the broader grant in
 * ~/.email-mcp/msal-cache.json under the same account, so the running MCP
 * server's email_calendar_event_create can then get a calendar token silently.
 * It does not read or write credentials.enc.
 *
 * Usage: node dist/cli/grant-calendar.js
 */
import { exec } from 'node:child_process';
import { OAuthCallbackServer } from '../auth/oauth-server.js';
import { OutlookAuth } from '../providers/outlook/auth.js';
import { OUTLOOK_CLIENT_ID } from '../oauth-config.js';

function openBrowser(url: string): void {
  const cmd =
    process.platform === 'darwin'
      ? `open "${url}"`
      : process.platform === 'win32'
        ? `start "" "${url}"`
        : `xdg-open "${url}"`;
  exec(cmd, (err) => {
    if (err) console.log(`\nCould not open a browser. Visit:\n${url}\n`);
  });
}

async function main(): Promise<void> {
  const server = new OAuthCallbackServer(300_000);
  const port = await server.start();
  const redirectUri = `http://localhost:${port}`;
  const scopes = [...OutlookAuth.SCOPES, ...OutlookAuth.CALENDAR_SCOPES];

  const auth = new OutlookAuth(OUTLOOK_CLIENT_ID);
  const { url, codeVerifier } = await auth.getAuthUrl(redirectUri, scopes);

  console.log('Sign in with the Outlook account that should create Teams meetings,');
  console.log('and accept the calendar permission. If no browser opens, visit:');
  console.log(url);
  openBrowser(url);

  let code: string;
  try {
    code = await server.waitForCode();
  } finally {
    server.shutdown();
  }

  const result = await auth.exchangeCode(code, codeVerifier, redirectUri, scopes);
  const account = result.account as { username?: string } | undefined;
  console.log(`Calendar access granted for ${account?.username ?? 'unknown account'}.`);
}

main().catch((err) => {
  console.error(`grant-calendar failed: ${err.message}`);
  process.exit(1);
});
