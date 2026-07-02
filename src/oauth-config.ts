// OAuth client settings.
// The Gmail client ID and secret are not kept in source. They come from env
// vars, or from a local JSON file only this Windows account can read:
//   EMAIL_MCP_GMAIL_CLIENT_ID + EMAIL_MCP_GMAIL_CLIENT_SECRET, else
//   EMAIL_MCP_GOOGLE_OAUTH_FILE, else
//   %LOCALAPPDATA%\VarietyVault\email-mcp\google-oauth.json  {"clientId","clientSecret"}
// The Outlook client ID is a public identifier, not a secret.
import { readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

function loadGoogleOAuth(): { clientId: string; clientSecret: string } {
  const envId = process.env.EMAIL_MCP_GMAIL_CLIENT_ID;
  const envSecret = process.env.EMAIL_MCP_GMAIL_CLIENT_SECRET;
  if (envId && envSecret) return { clientId: envId, clientSecret: envSecret };
  const file =
    process.env.EMAIL_MCP_GOOGLE_OAUTH_FILE ||
    join(process.env.LOCALAPPDATA || join(homedir(), '.config'), 'VarietyVault', 'email-mcp', 'google-oauth.json');
  try {
    const parsed = JSON.parse(readFileSync(file, 'utf8'));
    if (parsed.clientId && parsed.clientSecret) {
      return { clientId: String(parsed.clientId), clientSecret: String(parsed.clientSecret) };
    }
  } catch {
    // fall through to the warning below
  }
  console.error(`[email-mcp] Gmail OAuth settings missing: expected ${file} with clientId and clientSecret. Gmail accounts will fail to authenticate.`);
  return { clientId: '', clientSecret: '' };
}

const google = loadGoogleOAuth();
export const GMAIL_CLIENT_ID = google.clientId;
export const GMAIL_CLIENT_SECRET = google.clientSecret;
export const OUTLOOK_CLIENT_ID = '76d32225-ef10-4b49-951c-c80ae906595d';
