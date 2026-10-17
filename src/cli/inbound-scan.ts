/**
 * inbound-scan — unanswered-inbound detector + daily digest (BCC lane, v1).
 *
 * Standalone like watcher-fetch (no MCP server): loads ~/.email-mcp creds,
 * refreshes OAuth silently, scans every account for threads whose NEWEST
 * message is inbound with no later outbound, and emails Jon ONE digest.
 * Read-only against every mailbox except the single digest send to Jon
 * himself from the Vault-Shop account. Detection is deterministic — no LLM.
 *
 * IMPORTANT: credential store key derives from hostname+username, so this
 * MUST run as <user> (never SYSTEM). Scheduled via "BCC Inbound Digest".
 *
 * Usage:
 *   node dist/cli/inbound-scan.js [--dry-run] [--days 120]
 *     --dry-run  full scan, digest printed to stdout, nothing sent
 *
 * Output: run-summary lines to stdout (the launcher redirects to
 * business-command-center\logs\inbound_digest.log), report JSON to
 * business-command-center\data\inbound_scan_report.json, digest via email.
 * Exit 0 = scan completed and digest delivered (or printed in dry-run).
 */
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { dirname } from 'path';
import { AccountManager } from '../account-manager.js';
import type { EmailProvider } from '../providers/provider.js';
import type { Email } from '../models/types.js';

// Jon's own addresses — a message FROM any of these is outbound ("Jon
// replied"), regardless of which account it sits in. Includes the outreach
// alias, which is not an account. Source: email-action-sweep.js roster.
const OWN_ADDRESSES = new Set([
  'owner-1@example.invalid',
  'owner-2@example.invalid',
  'owner-3@example.invalid',
  'DIGEST_OWNER@example.invalid',
  'owner-6@example.invalid',
  'owner-7@example.invalid',
  'owner-8@example.invalid',
  'owner-4@example.invalid',
]);

const HANDLED_TAG = 'handled'; // Outlook category / Gmail label, case-insensitive
// Resolve the sender by ADDRESS, never by account id. Account ids are minted
// fresh by the setup wizard, so a re-auth silently invalidates a hardcoded id
// and the digest scans fine then dies at send (2026-08-12: the ENOSPC credential
// wipe forced a full re-auth and orphaned the old Vault-Shop id).
const DIGEST_FROM_EMAIL = 'DIGEST_OWNER@example.invalid';
const DIGEST_TO = 'DIGEST_OWNER@example.invalid';

const REGISTRY_PATH = process.env.GUAC_CUSTOMER_REGISTRY_PATH
  ?? '<shared-drive>\\order-intake-architecture\\catalog\\customer-registry.json';
const REPORT_PATH = process.env.INBOUND_SCAN_REPORT_PATH
  ?? 'C:\\Users\\<user>\\Documents\\business-command-center\\data\\inbound_scan_report.json';
// Written by dismiss_listener.py (the digest's "done" button, 127.0.0.1:8771).
const DISMISSALS_PATH = process.env.INBOUND_DISMISSALS_PATH
  ?? 'C:\\Users\\<user>\\Documents\\business-command-center\\data\\inbound_dismissals.json';
const DISMISS_URL = 'http://127.0.0.1:8771/dismiss';
// Jon-curated sender blacklist, taught via the triage board's spam button.
const SPAM_SENDERS_PATH = process.env.INBOUND_SPAM_SENDERS_PATH
  ?? 'C:\\Users\\<user>\\Documents\\business-command-center\\data\\inbound_spam_senders.json';

type Dismissals = Record<string, { last_seen_date?: string }>;

function loadDismissals(): Dismissals {
  try {
    return JSON.parse(readFileSync(DISMISSALS_PATH, 'utf-8'));
  } catch {
    return {};   // no file yet = nothing dismissed
  }
}

function loadSpamSenders(): Set<string> {
  try {
    return new Set(Object.keys(JSON.parse(readFileSync(SPAM_SENDERS_PATH, 'utf-8')))
      .map((s) => s.toLowerCase()));
  } catch {
    return new Set();   // no file yet = nothing spam-marked
  }
}

interface WaitingThread {
  accountName: string;
  threadId: string;
  subject: string;
  senderName: string;
  senderEmail: string;
  customer: string | null;       // registry display_name when resolved
  customerId: string | null;
  ageDays: number;
  newestInboundDate: string;
  snippet: string;
  messageCount: number;
  link: string | null;   // OWA webLink (Outlook) / rfc822msgid search (Gmail)
}

interface AccountHealth {
  name: string;
  email: string;
  ok: boolean;
  error?: string;
  scanned: number;
  waiting: number;
  droppedAutomated: number;
  droppedHandled: number;
}

function argValue(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

function log(msg: string): void {
  process.stdout.write(`${new Date().toISOString()} inbound-scan: ${msg}\n`);
}

// --- customer registry (mirror of spine/customer_registry.py resolve_email) ---

const PUBLIC_DOMAINS = new Set([
  'gmail.com', 'googlemail.com', 'aol.com', 'yahoo.com', 'outlook.com',
  'hotmail.com', 'live.com', 'icloud.com', 'me.com', 'msn.com',
]);

interface RegistryIndex {
  byEmail: Map<string, { customerId: string; displayName: string }>;
  byDomain: Map<string, { customerId: string; displayName: string }>;
  size: number;
}

function loadRegistry(): RegistryIndex {
  const idx: RegistryIndex = { byEmail: new Map(), byDomain: new Map(), size: 0 };
  try {
    const doc = JSON.parse(readFileSync(REGISTRY_PATH, 'utf-8'));
    for (const entry of doc.customers ?? []) {
      if (!entry?.customer_id || !entry?.display_name) continue;
      const rec = { customerId: entry.customer_id, displayName: entry.display_name };
      for (const e of entry.emails ?? []) {
        if (typeof e === 'string' && e.includes('@')) idx.byEmail.set(e.trim().toLowerCase(), rec);
      }
      for (const d of entry.domains ?? []) {
        if (typeof d === 'string' && d.trim() && !PUBLIC_DOMAINS.has(d.trim().toLowerCase())) {
          idx.byDomain.set(d.trim().toLowerCase(), rec);
        }
      }
      idx.size += 1;
    }
  } catch (err) {
    // Degrade to unresolved names — the digest still works without identity.
    log(`registry unreadable (${String(err)}), continuing without customer names`);
  }
  return idx;
}

function resolveCustomer(idx: RegistryIndex, addr: string) {
  const a = addr.trim().toLowerCase();
  return idx.byEmail.get(a) ?? idx.byDomain.get(a.split('@').pop() ?? '') ?? null;
}

// --- automated-sender filter ---

// no-reply variants match ANYWHERE in the local part (fbfs-no-reply@ was a
// live miss on the anchored version); the role-address list stays anchored.
const NOREPLY_ANYWHERE = /no-?_?reply|do-?not-?reply|donotreply/i;
const AUTOMATED_LOCALPART = /^(notifications?|notify|alerts?|mailer-daemon|postmaster|bounce[s]?|auto-?confirm|receipts?|billing|invoice\+|newsletter[s]?|news|marketing|promo(tions)?|offers|hello|digest|updates?|support|help|feedback|success|survey[s]?|webinars?|events?|insider|media|press|training|courses?|academy|community|membership|rewards)($|[@.+_-])/i;

// Gmail's own tab classification: anything Gmail filed as promotions, social,
// updates, or forums is not a customer waiting on a reply.
const GMAIL_BULK_LABELS = new Set([
  'CATEGORY_PROMOTIONS', 'CATEGORY_SOCIAL', 'CATEGORY_UPDATES', 'CATEGORY_FORUMS',
]);

const AUTOMATED_DOMAINS = [
  'amazonses.com', 'sendgrid.net', 'mailchimp.com', 'mailchimpapp.net',
  'mandrillapp.com', 'sparkpostmail.com', 'substack.com', 'shopifyemail.com',
  'shopify.com', 'stripe.com', 'intuit.com', 'quickbooks.com', 'ups.com',
  'fedex.com', 'usps.com', 'paypal.com', 'google.com', 'accounts.google.com',
  'youtube.com', 'microsoft.com', 'microsoftonline.com', 'github.com',
  'facebookmail.com', 'linkedin.com', 'x.com', 'twitter.com', 'ebay.com',
  'amazon.com', 'chase.com', 'wellsfargo.com', 'docusign.net', 'adobe.com',
  'zoom.us', 'calendly.com', 'godaddy.com', 'squarespace.com', 'wix.com',
  // live-run finds (dry-run 2, 2026-08-07): sponsored-blast + gov-newsletter
  // domains that Outlook left in Focused
  'aerospacemanufacturinganddesign.com', 'defenseandmunitions.com',
  'govcontractservice.com', 'senate.gov', 'fbfs.com',
  // 8/10 Jon flag: Nextdoor neighborhood digests ("Prairie Ridge Posts")
  'nextdoor.com',
];

// Bulk marketing pads snippets with invisible spacer characters (U+034F
// combining grapheme joiner, zero-width non-joiner/space). Humans never do.
// Kills the residual webinar/tradeshow blasts that dodge every other rule.
const INVISIBLE_MARKERS = /[͏​‌‍]{2,}|(?:[͏​‌‍]\s*){3,}/;

// Body-marker fallback for sponsored blasts whose sender domain we have not
// seen yet — the "This is a sponsored message" family all share footprints.
const SPONSORED_SNIPPET = /sponsored message|view (this email )?(online|as a web ?page)|email preferences|to unsubscribe/i;

function isAutomatedMessage(e: Email): boolean {
  const a = (e.from?.email ?? '').trim().toLowerCase();
  const [local = '', domain = ''] = [a.split('@')[0], a.split('@').pop() ?? ''];
  if (!domain.includes('.')) return true;   // root@localhost, bare hostnames
  if (local === 'email' || local === 'root') return true;
  if (NOREPLY_ANYWHERE.test(local)) return true;
  if (AUTOMATED_LOCALPART.test(local)) return true;
  if (SPONSORED_SNIPPET.test(e.snippet ?? '')) return true;
  if (INVISIBLE_MARKERS.test(e.snippet ?? '')) return true;
  if (AUTOMATED_DOMAINS.some((d) => domain === d || domain.endsWith(`.${d}`))) return true;
  // Platform-native bulk verdicts, strongest signals we have:
  if (e.inferenceClassification === 'other') return true;          // Outlook
  if ((e.labels ?? []).some((l) => GMAIL_BULK_LABELS.has(l))) return true;
  // RFC bulk-mail headers (populated on Gmail results; absent elsewhere):
  const h = (e as Email & { headers?: Record<string, string> }).headers;
  if (h) {
    if (h['list-unsubscribe'] || h['list-id'] || h['x-campaign'] || h['x-mailchimp-id']) return true;
    const prec = (h['precedence'] ?? '').toLowerCase();
    if (prec === 'bulk' || prec === 'list' || prec === 'junk') return true;
    if ((h['auto-submitted'] ?? '').toLowerCase().startsWith('auto')) return true;
  }
  return false;
}

// --- enumeration (provider-specific paging) ---

const PAGE = 100;

/** Outlook: server-side since-filter + top/skip paging (verified working). */
async function enumerateOutlook(provider: EmailProvider, folder: string,
                                sinceIso: string): Promise<Email[]> {
  const out: Email[] = [];
  for (let offset = 0; ; offset += PAGE) {
    const batch = await provider.search({
      folder, since: sinceIso, limit: PAGE, offset, returnBody: false,
    });
    out.push(...batch);
    if (batch.length < PAGE) break;
  }
  return out;
}

/** Gmail ignores `offset` entirely (no pageToken in the adapter), so page by
 *  date windows instead; a window that fills its limit splits in half so
 *  nothing is silently truncated. Folder is a Gmail label id (INBOX/SENT). */
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** One Gmail search with quota resilience: the adapter fetches every hit at
 *  format 'full' (5 quota units each), which can blow the per-minute user
 *  quota on a 120-day scan — first live run killed 2 of 5 accounts. On a
 *  quota error, wait for the minute window to refill and retry once. */
async function gmailSearchPaced(provider: EmailProvider, query: Parameters<EmailProvider['search']>[0]): Promise<Email[]> {
  try {
    const batch = await provider.search(query);
    await sleep(batch.length * 40);   // pace ~refill rate so we never spike
    return batch;
  } catch (err) {
    // quota = per-minute window refill; precondition/backend = Gmail transient
    // (killed the one Gmail account's scan in dry-run 2) — one paced retry for both.
    if (!/quota|precondition|rate ?limit|backend error|internal error/i.test(String(err))) throw err;
    await sleep(65_000);
    return provider.search(query);
  }
}

async function enumerateGmail(provider: EmailProvider, label: string,
                              since: Date, before: Date): Promise<Email[]> {
  const CAP = 200;
  const batch = await gmailSearchPaced(provider, {
    folder: label,
    since: since.toISOString(),
    before: before.toISOString(),
    limit: CAP,
    returnBody: false,
  });
  if (batch.length < CAP) return batch;
  const spanMs = before.getTime() - since.getTime();
  if (spanMs <= 2 * 24 * 3600 * 1000) {
    log(`WARNING: gmail window ${since.toISOString()}..${before.toISOString()} `
      + `hit the ${CAP} cap at minimum span - results truncated for ${label}`);
    return batch;
  }
  const mid = new Date(since.getTime() + spanMs / 2);
  const [a, b] = [
    await enumerateGmail(provider, label, since, mid),
    await enumerateGmail(provider, label, mid, before),
  ];
  return [...a, ...b];
}

async function enumerateAccount(provider: EmailProvider, providerType: string,
                                lookbackDays: number): Promise<Email[]> {
  const now = new Date();
  const since = new Date(now.getTime() - lookbackDays * 24 * 3600 * 1000);
  if (providerType === 'outlook') {
    const inbox = await enumerateOutlook(provider, 'inbox', since.toISOString());
    const sent = await enumerateOutlook(provider, 'sentitems', since.toISOString());
    return [...inbox, ...sent];
  }
  // gmail: windows of 15 days, split-on-cap
  const results: Email[] = [];
  for (const label of ['INBOX', 'SENT']) {
    let winStart = since;
    while (winStart < now) {
      const winEnd = new Date(Math.min(
        winStart.getTime() + 15 * 24 * 3600 * 1000, now.getTime() + 3600 * 1000));
      results.push(...await enumerateGmail(provider, label, winStart, winEnd));
      winStart = winEnd;
    }
  }
  return results;
}

// --- thread grouping + WAITING detection ---

function isOutbound(e: Email): boolean {
  return OWN_ADDRESSES.has((e.from?.email ?? '').trim().toLowerCase());
}

function hasHandledTag(e: Email): boolean {
  const tags = [...(e.categories ?? []), ...(e.labels ?? [])];
  return tags.some((t) => String(t).trim().toLowerCase() === HANDLED_TAG);
}

interface ScanCounters { droppedAutomated: number; droppedHandled: number }

function findWaitingThreads(accountName: string, messages: Email[],
                            registry: RegistryIndex, counters: ScanCounters,
                            now: Date, dismissals: Dismissals = {},
                            spamSenders: Set<string> = new Set()): WaitingThread[] {
  // Group by threadId; a message with no threadId is its own thread.
  const threads = new Map<string, Email[]>();
  for (const m of messages) {
    if (m.flags?.draft) continue;                     // drafts are not traffic
    const key = m.threadId || `solo:${m.id}`;
    const list = threads.get(key);
    if (list) {
      if (!list.some((x) => x.id === m.id)) list.push(m); // inbox+sent overlap
    } else {
      threads.set(key, [m]);
    }
  }

  const waiting: WaitingThread[] = [];
  for (const [threadId, msgs] of threads) {
    msgs.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const newest = msgs[msgs.length - 1];
    if (isOutbound(newest)) continue;                 // Jon spoke last - fine
    const sender = (newest.from?.email ?? '').trim().toLowerCase();
    if (!sender) continue;
    if (isAutomatedMessage(newest)) {
      counters.droppedAutomated += 1;
      continue;
    }
    if (hasHandledTag(newest)) {                      // Jon's dismiss button
      counters.droppedHandled += 1;
      continue;
    }
    // Board spam button: Jon-taught sender blacklist, permanent (unlike a
    // dismissal, new mail from these senders never resurfaces).
    if (spamSenders.has((newest.from?.email ?? '').toLowerCase())) {
      counters.droppedAutomated += 1;
      continue;
    }
    // Digest "done" button: dismissal covers the thread AS OF the message that
    // was newest at click time - anything newer resurfaces it.
    const dis = dismissals[threadId];
    if (dis?.last_seen_date
        && new Date(newest.date).getTime() <= Date.parse(dis.last_seen_date)) {
      counters.droppedHandled += 1;
      continue;
    }
    const ageMs = now.getTime() - new Date(newest.date).getTime();
    const customer = resolveCustomer(registry, sender);
    const headers = (newest as Email & { headers?: Record<string, string> }).headers;
    const msgId = headers?.['message-id']?.replace(/^<|>$/g, '');
    const link = newest.webLink
      ?? (msgId ? `https://mail.google.com/mail/#search/rfc822msgid:${encodeURIComponent(msgId)}` : null);
    waiting.push({
      accountName,
      threadId,
      subject: newest.subject || '(no subject)',
      senderName: newest.from?.name || sender,
      senderEmail: sender,
      customer: customer?.displayName ?? null,
      customerId: customer?.customerId ?? null,
      ageDays: Math.max(0, Math.floor(ageMs / (24 * 3600 * 1000))),
      newestInboundDate: newest.date,
      snippet: (newest.snippet ?? '').replace(/\s+/g, ' ').slice(0, 160),
      messageCount: msgs.length,
      link,
    });
  }
  return waiting;
}

// --- digest rendering (plain text, Jon's rules: no em dashes, short) ---

function renderDigest(waiting: WaitingThread[], health: AccountHealth[],
                      now: Date,
                      lookbackDays: number): { subject: string; text: string } {
  const day = now.toISOString().slice(0, 10);
  const subject = waiting.length === 0
    ? `Inbound Digest ${day} - all clear`
    : `Inbound Digest ${day} - ${waiting.length} waiting`;

  const lines: string[] = [];
  lines.push('MAILBOX HEALTH');
  for (const h of health) {
    lines.push(h.ok
      ? `  OK    ${h.name} (${h.email}) - ${h.scanned} msgs scanned, ${h.waiting} waiting`
      : `  RED   ${h.name} (${h.email}) - ${h.error ?? 'unreachable'}`);
  }
  lines.push('');

  const OTHERS_MAX = 40;   // full list always lands in the report JSON
  const customers = waiting.filter((w) => w.customer !== null);
  const others = waiting.filter((w) => w.customer === null);

  const item = (w: WaitingThread, withLink: boolean) => {
    const who = w.customer ? `${w.customer} (${w.senderName})` : w.senderName;
    lines.push(`${String(w.ageDays).padStart(3)}d  ${who} <${w.senderEmail}>`);
    lines.push(`      ${w.subject}  [${w.accountName}]`);
    if (w.snippet) lines.push(`      "${w.snippet}"`);
    if (withLink && w.link) lines.push(`      open: ${w.link}`);
    lines.push('');
  };

  if (waiting.length === 0) {
    lines.push('All clear. No customer threads are waiting on a reply.');
  } else {
    if (customers.length > 0) {
      lines.push(`CUSTOMERS WAITING (${customers.length}, oldest first)`);
      lines.push('');
      customers.forEach((w) => item(w, true));
    }
    if (others.length > 0) {
      lines.push(`OTHERS - unrecognized senders (${others.length}, oldest first`
        + (others.length > OTHERS_MAX ? `, showing ${OTHERS_MAX}` : '') + ')');
      lines.push('');
      others.slice(0, OTHERS_MAX).forEach((w) => item(w, false));
      if (others.length > OTHERS_MAX) {
        lines.push(`...and ${others.length - OTHERS_MAX} more, full list in `
          + `data\\inbound_scan_report.json`);
        lines.push('');
      }
    }
    lines.push('Answered by text or phone? Tag the thread "Handled" in Outlook');
    lines.push('or Gmail and it drops off tomorrow. A reply clears it automatically.');
  }
  lines.push('');
  lines.push(`Scanned ${health.reduce((n, h) => n + h.scanned, 0)} messages across `
    + `${health.filter((h) => h.ok).length}/${health.length} accounts, ${lookbackDays}-day window.`);
  lines.push('This digest always sends. A morning without one means the scanner broke.');
  return { subject, text: lines.join('\n') };
}

// --- HTML digest (email-client-safe: tables + inline styles only) ---

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function ageBadge(days: number): string {
  const bg = days >= 14 ? '#8A2020' : days >= 4 ? '#B4541A' : '#55524B';
  return `<td width="52" align="center" valign="top" style="padding:10px 0 0 0">`
    + `<div style="background:${bg};color:#ffffff;border-radius:6px;font:bold 13px Arial;`
    + `padding:6px 0;width:46px">${days}d</div></td>`;
}

function renderDigestHtml(waiting: WaitingThread[], health: AccountHealth[],
                          now: Date, lookbackDays: number): string {
  const day = now.toISOString().slice(0, 10);
  const OTHERS_MAX = 40;
  const customers = waiting.filter((w) => w.customer !== null);
  const others = waiting.filter((w) => w.customer === null);
  const P = '#582960'; // Vault Purple

  const healthRows = health.map((h) => {
    const dot = h.ok ? '#2E6B3E' : '#8A2020';
    const detail = h.ok
      ? `${h.scanned} scanned &middot; <b>${h.waiting} waiting</b> &middot; ${h.droppedAutomated} auto-dropped`
      : esc(h.error ?? 'unreachable');
    return `<tr><td width="14" style="font-size:13px;color:${dot};padding:2px 6px 2px 0">&#9679;</td>`
      + `<td style="font:13px Arial;color:#26221E;padding:2px 0"><b>${esc(h.name)}</b>`
      + ` <span style="color:#6A6670">${detail}</span></td></tr>`;
  }).join('');

  const card = (w: WaitingThread, withLink: boolean) => {
    const who = w.customer
      ? `<span style="color:${P}">${esc(w.customer)}</span> <span style="color:#6A6670;font-weight:normal">(${esc(w.senderName)})</span>`
      : esc(w.senderName);
    const open = withLink && w.link
      ? ` &nbsp;<a href="${esc(w.link)}" style="color:${P};font:bold 12px Arial;text-decoration:none;border:1px solid ${P};border-radius:4px;padding:2px 8px">open</a>`
      : '';
    const done = ` &nbsp;<a href="${DISMISS_URL}?tid=${encodeURIComponent(w.threadId)}`
      + `&d=${encodeURIComponent(w.newestInboundDate)}`
      + `&who=${encodeURIComponent(w.customer ?? w.senderName)}"`
      + ` style="color:#2E6B3E;font:bold 12px Arial;text-decoration:none;`
      + `border:1px solid #2E6B3E;border-radius:4px;padding:2px 8px">&#10003; done</a>`;
    return `<tr><td colspan="2" style="border-top:1px solid #E4E0DA;font-size:1px;line-height:1px">&nbsp;</td></tr>`
      + `<tr>${ageBadge(w.ageDays)}<td style="padding:8px 0 10px 10px">`
      + `<div style="font:bold 14px Arial;color:#26221E">${who}${open}${done}</div>`
      + `<div style="font:13px Arial;color:#26221E;padding-top:2px">${esc(w.subject)}`
      + ` <span style="color:#9B968F">[${esc(w.accountName)}]</span></div>`
      + (w.snippet ? `<div style="font:italic 12px Arial;color:#6A6670;padding-top:3px">&ldquo;${esc(w.snippet)}&rdquo;</div>` : '')
      + `</td></tr>`;
  };

  const section = (title: string, count: string) =>
    `<tr><td colspan="2" style="padding:18px 0 6px 0">`
    + `<div style="font:bold 11px Arial;letter-spacing:2px;color:#ffffff;background:${P};`
    + `border-radius:4px;padding:5px 10px;display:inline-block">${title}`
    + ` &nbsp;${count}</div></td></tr>`;

  let body = '';
  if (waiting.length === 0) {
    body = `<tr><td style="font:15px Arial;color:#2E6B3E;padding:16px 0"><b>All clear.</b> No customer threads are waiting on a reply.</td></tr>`;
  } else {
    if (customers.length > 0) {
      body += section('CUSTOMERS WAITING', `${customers.length}, oldest first`);
      customers.forEach((w) => { body += card(w, true); });
    }
    if (others.length > 0) {
      body += section('OTHERS &middot; UNRECOGNIZED SENDERS',
        `${others.length}${others.length > OTHERS_MAX ? `, showing ${OTHERS_MAX}` : ''}`);
      others.slice(0, OTHERS_MAX).forEach((w) => { body += card(w, false); });
      if (others.length > OTHERS_MAX) {
        body += `<tr><td colspan="2" style="font:12px Arial;color:#6A6670;padding:8px 0">`
          + `&hellip;and ${others.length - OTHERS_MAX} more in data\\inbound_scan_report.json</td></tr>`;
      }
    }
  }

  const scanned = health.reduce((n, h) => n + h.scanned, 0);
  const okCount = health.filter((h) => h.ok).length;
  return `<div style="background:#F4F2EF;padding:18px 0">`
    + `<table cellpadding="0" cellspacing="0" border="0" align="center" width="680" style="max-width:680px;background:#ffffff;border:1px solid #E0DCD6;border-radius:8px">`
    + `<tr><td style="background:#26221E;border-radius:8px 8px 0 0;padding:14px 22px">`
    + `<span style="font:bold 17px Arial;color:#ffffff;letter-spacing:1px">INBOUND DIGEST</span>`
    + `<span style="font:13px Arial;color:#C9C4BD">&nbsp;&nbsp;${day}</span>`
    + `<span style="float:right;font:bold 14px Arial;color:#ffffff;background:${waiting.length ? P : '#2E6B3E'};border-radius:12px;padding:3px 12px">${waiting.length ? `${waiting.length} waiting` : 'all clear'}</span>`
    + `</td></tr>`
    // Triage board strip: loopback URL, so it only resolves on Jon's own PC
    // (same reach as the dismiss buttons; on the phone it simply won't load).
    + `<tr><td style="background:#582960;padding:8px 22px">`
    + `<a href="http://127.0.0.1:8771/board" style="font:bold 12px Arial;color:#ffffff;text-decoration:none">&#9654; open the full triage board &mdash; every item, done buttons, and a note box that reaches guac-bcc</a>`
    + `</td></tr>`
    + `<tr><td style="padding:14px 22px 4px 22px">`
    + `<table cellpadding="0" cellspacing="0" border="0" width="100%">`
    + `<tr><td colspan="2" style="font:bold 11px Arial;letter-spacing:2px;color:#6A6670;padding-bottom:4px">MAILBOX HEALTH</td></tr>`
    + healthRows
    + `</table></td></tr>`
    + `<tr><td style="padding:0 22px 8px 22px"><table cellpadding="0" cellspacing="0" border="0" width="100%">${body}</table></td></tr>`
    + `<tr><td style="background:#F4F2EF;border-radius:0 0 8px 8px;border-top:1px solid #E0DCD6;padding:12px 22px">`
    + `<div style="font:12px Arial;color:#6A6670">Answered by text or phone? Tag the thread <b>Handled</b> in Outlook or Gmail and it drops off tomorrow. A reply clears it automatically.</div>`
    + `<div style="font:11px Arial;color:#9B968F;padding-top:6px">Scanned ${scanned} messages across ${okCount}/${health.length} accounts, ${lookbackDays}-day window &middot; this digest always sends; a morning without one means the scanner broke.</div>`
    + `</td></tr></table></div>`;
}

// --- main ---

/** Declared cadence: daily at 06:57 local (task "BCC Inbound Digest").
 *  Carried in the report so consumers judge staleness against producer
 *  intent instead of a hardcoded guess (mesh feed convention). */
function nextExpectedRun(now: Date): string {
  const next = new Date(now);
  next.setHours(6, 57, 0, 0);
  if (next <= now) next.setDate(next.getDate() + 1);
  return next.toISOString();
}

async function main(): Promise<void> {
  const dryRun = process.argv.includes('--dry-run');
  const renderOnly = process.argv.includes('--render-only');
  const lookbackDays = parseInt(argValue('--days') ?? '120', 10);
  const now = new Date();

  // --render-only: re-send the digest from the LAST report on disk, no scan.
  // For iterating on the visual format without a 30-minute mailbox sweep.
  // Look the digest sender up by address. Throws loudly if the mailbox is not
  // authenticated, which is far better than silently scanning and never sending.
  async function digestSender(m: AccountManager) {
    const all = await m.listAccounts();
    const hit = all.find((a) => a.email.toLowerCase() === DIGEST_FROM_EMAIL);
    if (!hit) {
      throw new Error(`digest sender ${DIGEST_FROM_EMAIL} is not an authenticated `
        + `account (have: ${all.map((a) => a.email).join(', ')}). Re-add it with `
        + `npm run setup.`);
    }
    return m.getProvider(hit.id);
  }

  if (renderOnly) {
    const doc = JSON.parse(readFileSync(REPORT_PATH, 'utf-8'));
    const waiting: WaitingThread[] = doc.waiting ?? [];
    const health: AccountHealth[] = doc.health ?? [];
    const digest = renderDigest(waiting, health, now, doc.lookback_days ?? lookbackDays);
    const html = renderDigestHtml(waiting, health, now, doc.lookback_days ?? lookbackDays);
    const mgrRo = new AccountManager();
    const sender = await digestSender(mgrRo);
    const sent = await sender.sendEmail({
      to: [{ name: 'Owner', email: DIGEST_TO }],
      subject: `${digest.subject} - visual preview`,
      body: { text: digest.text, html },
    });
    log(`render-only digest sent from report of ${doc.generated_at} (id ${sent.id})`);
    await mgrRo.disconnectAll();
    return;
  }

  const registry = loadRegistry();
  const dismissals = loadDismissals();
  const spamSenders = loadSpamSenders();
  log(`start dry-run=${dryRun} lookback=${lookbackDays}d registry=${registry.size} `
    + `customers, ${Object.keys(dismissals).length} dismissals, `
    + `${spamSenders.size} spam senders on file`);

  const mgr = new AccountManager();
  const accounts = await mgr.listAccounts();
  const health: AccountHealth[] = [];
  const allWaiting: WaitingThread[] = [];

  for (const acct of accounts) {
    const h: AccountHealth = {
      name: acct.name, email: acct.email, ok: false,
      scanned: 0, waiting: 0, droppedAutomated: 0, droppedHandled: 0,
    };
    health.push(h);
    // One full-account retry after a cool-down: dry-runs 2 and 3 each lost a
    // DIFFERENT Gmail account to transient API errors (quota, precondition,
    // invalid_request) that succeeded on the next run. Second failure = RED.
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        const provider = await mgr.getProvider(acct.id);
        const test = await provider.testConnection();
        if (!test.success) throw new Error(test.error ?? 'connection test failed');
        const messages = await enumerateAccount(provider, acct.provider, lookbackDays);
        const counters: ScanCounters = { droppedAutomated: 0, droppedHandled: 0 };
        const waiting = findWaitingThreads(acct.name, messages, registry, counters, now, dismissals, spamSenders);
        h.ok = true;
        h.error = undefined;
        h.scanned = messages.length;
        h.waiting = waiting.length;
        h.droppedAutomated = counters.droppedAutomated;
        h.droppedHandled = counters.droppedHandled;
        allWaiting.push(...waiting);
        log(`${acct.name}: ${messages.length} msgs, ${waiting.length} waiting, `
          + `${counters.droppedAutomated} automated dropped, ${counters.droppedHandled} handled`);
        break;
      } catch (err) {
        h.error = err instanceof Error ? err.message : String(err);
        if (attempt === 1 && !/invalid_grant/i.test(h.error)) {
          log(`${acct.name}: attempt 1 failed (${h.error}), retrying in 70s`);
          await sleep(70_000);
        } else {
          log(`${acct.name}: RED - ${h.error}`);
          break;
        }
      }
    }
  }

  allWaiting.sort((a, b) => b.ageDays - a.ageDays);   // oldest first

  try {
    mkdirSync(dirname(REPORT_PATH), { recursive: true });
    writeFileSync(REPORT_PATH, JSON.stringify({
      generated_at: now.toISOString(),
      next_expected_at: nextExpectedRun(now),
      lookback_days: lookbackDays,
      registry_customers: registry.size,
      health,
      waiting: allWaiting,
    }, null, 2));
    log(`report written: ${REPORT_PATH}`);
  } catch (err) {
    log(`WARNING: report write failed (${String(err)}), digest continues`);
  }

  const digest = renderDigest(allWaiting, health, now, lookbackDays);
  if (dryRun) {
    process.stdout.write(`\n----- DIGEST PREVIEW (not sent) -----\n`
      + `Subject: ${digest.subject}\n\n${digest.text}\n`
      + `----- END PREVIEW -----\n`);
  } else {
    const sender = await digestSender(mgr);
    const sent = await sender.sendEmail({
      to: [{ name: 'Owner', email: DIGEST_TO }],
      subject: digest.subject,
      body: { text: digest.text, html: renderDigestHtml(allWaiting, health, now, lookbackDays) },
    });
    log(`digest sent: ${digest.subject} (id ${sent.id})`);
  }

  await mgr.disconnectAll();   // IMAP/Graph holds the loop open otherwise
  log(`done: ${allWaiting.length} waiting across `
    + `${health.filter((h) => h.ok).length}/${health.length} accounts`);
}

main().catch(async (err) => {
  process.stderr.write(`inbound-scan: ${err?.stack ?? String(err)}\n`);
  process.exit(1);
});
