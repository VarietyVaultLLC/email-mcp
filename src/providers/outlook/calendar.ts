import { Client } from '@microsoft/microsoft-graph-client';

export interface CreateEventParams {
  subject: string;
  /** Local wall-clock start, e.g. "2026-10-01T14:00:00", interpreted in timeZone. */
  start: string;
  end: string;
  /** Windows or IANA zone name Graph accepts, e.g. "Central Standard Time". */
  timeZone: string;
  attendees?: Array<{ email: string; name?: string }>;
  body?: { text?: string; html?: string };
  location?: string;
  teamsMeeting: boolean;
}

export interface CreatedEvent {
  id: string;
  webLink: string | null;
  joinUrl: string | null;
  attendeesInvited: string[];
}

/**
 * Build the Graph `POST /me/events` body.
 *
 * isOnlineMeeting + onlineMeetingProvider "teamsForBusiness" makes Exchange
 * create the Teams meeting and put the join link in the invite body.
 */
export function buildEventPayload(p: CreateEventParams): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    subject: p.subject,
    start: { dateTime: p.start, timeZone: p.timeZone },
    end: { dateTime: p.end, timeZone: p.timeZone },
    isOnlineMeeting: p.teamsMeeting,
  };
  if (p.teamsMeeting) payload.onlineMeetingProvider = 'teamsForBusiness';
  if (p.body?.html) payload.body = { contentType: 'HTML', content: p.body.html };
  else if (p.body?.text) payload.body = { contentType: 'Text', content: p.body.text };
  if (p.location) payload.location = { displayName: p.location };
  if (p.attendees?.length) {
    payload.attendees = p.attendees.map((a) => ({
      emailAddress: { address: a.email, ...(a.name ? { name: a.name } : {}) },
      type: 'required',
    }));
  }
  return payload;
}

/**
 * Create the event on the signed-in user's calendar. With attendees, Exchange
 * sends the invitations immediately; there is no draft state for events.
 */
export async function createOutlookEvent(
  accessToken: string,
  p: CreateEventParams
): Promise<CreatedEvent> {
  const client = Client.init({ authProvider: (done) => done(null, accessToken) });
  const created = await client.api('/me/events').post(buildEventPayload(p));
  return {
    id: created.id,
    webLink: created.webLink ?? null,
    joinUrl: created.onlineMeeting?.joinUrl ?? null,
    attendeesInvited: (p.attendees ?? []).map((a) => a.email),
  };
}
