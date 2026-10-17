import { describe, it, expect } from 'vitest';
import { buildEventPayload } from '../../src/providers/outlook/calendar.js';
import { OutlookAuth } from '../../src/providers/outlook/auth.js';

const base = {
  subject: 'Project EMU follow-up',
  start: '2026-10-01T14:00:00',
  end: '2026-10-01T14:30:00',
  timeZone: 'Central Standard Time',
};

describe('buildEventPayload', () => {
  it('asks Exchange for a Teams meeting by default', () => {
    const p = buildEventPayload({ ...base, teamsMeeting: true });
    expect(p.isOnlineMeeting).toBe(true);
    expect(p.onlineMeetingProvider).toBe('teamsForBusiness');
    expect(p.start).toEqual({ dateTime: base.start, timeZone: base.timeZone });
  });

  it('omits the provider when no Teams meeting is wanted', () => {
    const p = buildEventPayload({ ...base, teamsMeeting: false });
    expect(p.isOnlineMeeting).toBe(false);
    expect(p).not.toHaveProperty('onlineMeetingProvider');
  });

  it('adds no attendees key when none are given, so no invites go out', () => {
    const p = buildEventPayload({ ...base, teamsMeeting: true });
    expect(p).not.toHaveProperty('attendees');
  });

  it('maps attendees as required and prefers html body over text', () => {
    const p = buildEventPayload({
      ...base,
      teamsMeeting: true,
      attendees: [{ email: 'a@example.com', name: 'A' }, { email: 'b@example.com' }],
      body: { text: 'plain', html: '<p>rich</p>' },
    });
    expect(p.attendees).toEqual([
      { emailAddress: { address: 'a@example.com', name: 'A' }, type: 'required' },
      { emailAddress: { address: 'b@example.com' }, type: 'required' },
    ]);
    expect(p.body).toEqual({ contentType: 'HTML', content: '<p>rich</p>' });
  });
});

describe('OutlookAuth scopes', () => {
  it('keeps calendar out of the mail scopes (incremental consent)', () => {
    expect(OutlookAuth.SCOPES).not.toContain('Calendars.ReadWrite');
    expect(OutlookAuth.CALENDAR_SCOPES).toEqual(['Calendars.ReadWrite']);
  });
});
