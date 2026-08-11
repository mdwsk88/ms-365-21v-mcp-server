import assert from 'node:assert/strict';
import test from 'node:test';

import { buildCalendarInsights, buildMailDigest, buildTeamsDigest } from '../dist/tools/smart.js';

test('mail digest groups internal, external, unread, flagged, and attachment signals', () => {
  const digest = buildMailDigest(
    [
      {
        id: '1',
        subject: 'Internal urgent',
        from: { emailAddress: { address: 'boss@company.cn' } },
        receivedDateTime: '2026-07-15T08:00:00Z',
        isRead: false,
        importance: 'high',
        hasAttachments: true,
        flag: { flagStatus: 'flagged' }
      },
      {
        id: '2',
        subject: 'External update',
        from: { emailAddress: { address: 'vendor@example.com' } },
        receivedDateTime: '2026-07-15T07:00:00Z',
        isRead: false,
        importance: 'normal',
        hasAttachments: false,
        flag: { flagStatus: 'notFlagged' }
      },
      {
        id: '3',
        subject: 'Read mail',
        from: { emailAddress: { address: 'person@company.cn' } },
        receivedDateTime: '2026-07-15T06:00:00Z',
        isRead: true,
        importance: 'normal',
        hasAttachments: false
      }
    ],
    'company.cn',
    'unread'
  );
  assert.deepEqual(digest.summary, {
    total: 3,
    unread: 2,
    flagged: 1,
    withAttachments: 1,
    internal: 2,
    external: 1
  });
  assert.equal(digest.highlights.length, 2);
  assert.equal(digest.highlights[0].subject, 'Internal urgent');
  assert.equal(digest.byCategory.external_unread.length, 1);
});

test('calendar insights detect overlap and calculate free working-hour slots', () => {
  const insights = buildCalendarInsights(
    [
      {
        id: 'event-1',
        subject: 'First',
        start: { dateTime: '2026-07-15T09:00:00', timeZone: 'China Standard Time' },
        end: { dateTime: '2026-07-15T10:30:00', timeZone: 'China Standard Time' }
      },
      {
        id: 'event-2',
        subject: 'Second',
        start: { dateTime: '2026-07-15T10:00:00', timeZone: 'China Standard Time' },
        end: { dateTime: '2026-07-15T11:00:00', timeZone: 'China Standard Time' }
      }
    ],
    1,
    new Date('2026-07-15T00:00:00Z')
  );
  assert.equal(insights.totalEvents, 2);
  assert.equal(insights.conflicts.length, 1);
  assert.equal(insights.conflicts[0].overlapMinutes, 30);
  assert.equal(insights.busyRatio, 0.22);
  assert.match(insights.conflicts[0].suggestion, /11:00-18:00/);
  assert.ok(insights.freeSlots.some((slot) => slot.start === '11:00' && slot.timeZone === 'China Standard Time'));
});

test('Teams digest labels recent messages as candidates instead of claiming exact unread state', () => {
  const digest = buildTeamsDigest(
    [
      {
        sourceType: 'channel',
        sourceName: 'Ops / General',
        messages: [
          {
            id: 'message-1',
            createdDateTime: '2026-07-15T08:00:00Z',
            from: { user: { displayName: 'Alice' } },
            body: { content: '<b>Attention</b> deployment changed' },
            mentions: [{ id: 1 }],
            attachments: []
          }
        ]
      },
      {
        sourceType: 'chat',
        sourceName: 'Project chat',
        messages: [
          {
            id: 'message-2',
            createdDateTime: '2026-07-15T09:00:00Z',
            from: { user: { displayName: 'Bob' } },
            body: { content: 'Normal update' },
            mentions: [],
            attachments: []
          }
        ]
      }
    ],
    20
  );
  assert.equal(digest.unreadStatusAvailable, false);
  assert.equal(digest.items[0].messageId, 'message-1');
  assert.equal(digest.items[0].bodyPreview, 'Attention deployment changed');
});
