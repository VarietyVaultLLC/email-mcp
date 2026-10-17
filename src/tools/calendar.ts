import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { AccountManager } from '../account-manager.js';
import { createOutlookEvent } from '../providers/outlook/calendar.js';

function jsonResult(data: unknown) {
  return { content: [{ type: 'text' as const, text: JSON.stringify(data) }] };
}

export function registerCalendarTools(server: McpServer, accountManager: AccountManager): void {
  // --- email_calendar_event_create ---
  server.tool(
    'email_calendar_event_create',
    'Create an Outlook calendar event, with a Microsoft Teams meeting link by default (Outlook only). ' +
      'WITH attendees, Outlook sends the invitations IMMEDIATELY, like email_send: get the user\'s explicit approval first. ' +
      'WITHOUT attendees, it only creates the meeting on the organizer\'s calendar and returns the Teams joinUrl.',
    {
      accountId: z.string(),
      subject: z.string(),
      start: z.string().describe('Local start time, e.g. 2026-10-01T14:00:00'),
      end: z.string().describe('Local end time, e.g. 2026-10-01T14:30:00'),
      timeZone: z.string().optional().describe('Default "Central Standard Time"'),
      attendees: z.array(z.object({ email: z.string(), name: z.string().optional() })).optional(),
      body: z.object({ text: z.string().optional(), html: z.string().optional() }).optional(),
      location: z.string().optional(),
      teamsMeeting: z.boolean().optional().describe('Default true'),
    },
    async (args) => {
      try {
        if (Date.parse(args.end) <= Date.parse(args.start)) {
          return jsonResult({ success: false, error: 'end must be after start' });
        }
        const token = await accountManager.getOutlookCalendarToken(args.accountId);
        const created = await createOutlookEvent(token, {
          subject: args.subject,
          start: args.start,
          end: args.end,
          timeZone: args.timeZone ?? 'Central Standard Time',
          attendees: args.attendees,
          body: args.body,
          location: args.location,
          teamsMeeting: args.teamsMeeting ?? true,
        });
        const teamsLinkMissing = (args.teamsMeeting ?? true) && !created.joinUrl;
        return jsonResult({
          success: true,
          data: created,
          ...(teamsLinkMissing
            ? { warning: 'Event created but Graph returned no Teams joinUrl; check the Teams meeting policy (Outlook add-in) for this user.' }
            : {}),
        });
      } catch (error: any) {
        return jsonResult({ success: false, error: error.message });
      }
    }
  );
}
