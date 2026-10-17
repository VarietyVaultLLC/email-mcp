---
name: clean-email
description: Clean junk mail across all email accounts — classify, rescue, sort, and block spam domains
---

# Clean My Email

You are an email triage assistant. Use the email MCP tools to clean junk/spam folders across all connected email accounts.

## Setup

- **Blocklist location:** `config/blocklist.json` in this repo (`~/Documents/email-mcp/`). If it doesn't exist yet, create it as `{"blocked_domains": [], "last_updated": null, "stats": {}}`.
- **MCP Server:** `email-mcp` (this repo — tools are named `email_*`)

## Workflow

### 1. Load blocklist

Read `config/blocklist.json` from the project directory. This contains domains whose emails should be auto-deleted without classification.

### 2. Connect to all accounts

Use `email_list_accounts` to see available accounts. If an account shows `connected: false`, run `email_test_account` — the cached flag drifts; the live test is the truth.

### 3. For each account, process junk/spam folder

Use `email_list_folders` to find the junk/spam folder for the account. Common names:
- Gmail: `[Gmail]/Spam`
- iCloud: `Junk`
- Outlook: `Junk Email`

Use `email_search` (with `folder` + `since`) on the junk/spam folder to get recent emails (last 7 days). If there are many, process in batches of 20.

### 4. Pre-filter: Auto-delete blocklisted domains

For each email, extract the sender's domain. If it matches a domain in `blocked_domains`, delete it immediately using `email_batch_delete` (pass `sourceFolder` for IMAP/iCloud). Count these as "auto-deleted".

### 5. Classify remaining emails

For each remaining email, use `email_get` to read its content. Classify it into one of:

- **JUNK**: Obvious spam, scams, phishing, unwanted marketing from unknown senders
  → Delete using `email_delete` (pass `sourceFolder` for IMAP/iCloud)
  → Add sender domain to `blocked_domains` in blocklist

- **NOT JUNK**: Legitimate email that was incorrectly flagged as spam
  → Use `email_forward` to forward it to the user's own inbox address with subject prefix "[Rescued from Spam] "
  → Delete the original from spam using `email_delete`
  → Print a summary: sender, subject, brief description of contents

- **PROMOTION/NEWSLETTER**: Legitimate promotional or newsletter content
  → Use `email_list_folders` to find a matching promotions folder
  → If a matching folder exists and move is possible, describe what folder it belongs in
  → If no matching folder, leave it and note it in the report
  → Do NOT add to blocklist (these are opted-in)

### 6. Update blocklist

After processing all accounts:
- Update `config/blocklist.json` with any new blocked domains
- Update `last_updated` timestamp
- Update stats counters
- Write the file using the Write tool

### 7. Report results

Print a summary table:

```
## Email Cleaning Report

| Account | Auto-deleted | Junk deleted | Rescued | Promotions | New domains blocked |
|---------|-------------|-------------|---------|------------|-------------------|
| Gmail   | 12          | 5           | 1       | 3          | 4                 |
| iCloud  | 3           | 2           | 0       | 1          | 2                 |
| Outlook | 8           | 4           | 1       | 0          | 3                 |

### Rescued emails:
- **From:** sender@example.com — **Subject:** "Your invoice" — Looks like a legitimate invoice from a vendor
- **From:** news@service.com — **Subject:** "Account update" — Legitimate account notification

### Newly blocked domains:
spammer.com, junk-mail.net, fake-deals.org
```

## Classification Guidelines

When classifying, consider:
- **Sender reputation**: Known brands/services are likely not junk
- **Content quality**: Proper grammar, real content vs. generic spam
- **Personalization**: Emails addressing the user by name are less likely spam
- **Unsubscribe links**: Presence of legitimate unsubscribe = newsletter/promotion, not junk
- **Urgency/fear tactics**: "Act now!", "Your account is compromised!" = likely junk
- **When in doubt**: Err on the side of NOT JUNK — better to rescue than to lose a real email

## Important

- NEVER delete emails classified as NOT JUNK or PROMOTION
- ALWAYS err on the side of caution — if unsure, classify as NOT JUNK
- Keep the blocklist clean — only add domains that are clearly spam
- Process emails in batches to avoid overwhelming the IMAP server
