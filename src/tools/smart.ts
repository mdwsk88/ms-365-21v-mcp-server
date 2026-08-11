import { z } from 'zod/v4';
import type { AppConfig } from '../config.js';
import {
  getMe,
  listChannelMessages,
  listChatMessages,
  listMyCalendarEvents,
  listMyChats,
  listMyJoinedTeams,
  listMyMessagesSince,
  listTeamChannels
} from '../graph.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

type JsonRecord = Record<string, unknown>;

function record(value: unknown): JsonRecord {
  return value && typeof value === 'object' ? (value as JsonRecord) : {};
}

function recordsFromCollection(value: unknown): JsonRecord[] {
  const items = record(value).value;
  return Array.isArray(items) ? items.map(record) : [];
}

function stringValue(value: unknown): string | undefined {
  return typeof value === 'string' && value ? value : undefined;
}

function emailAddress(value: unknown): string | undefined {
  return stringValue(record(record(value).emailAddress).address);
}

function emailDomain(value: string | undefined): string | undefined {
  const at = value?.lastIndexOf('@') ?? -1;
  return at >= 0 ? value?.slice(at + 1).toLowerCase() : undefined;
}

type MailFocus = 'all' | 'unread' | 'flagged';

export function buildMailDigest(messages: JsonRecord[], internalDomain: string | undefined, focus: MailFocus) {
  const items = messages.map((message) => {
    const sender = emailAddress(message.from) ?? 'unknown';
    const flagged = stringValue(record(message.flag).flagStatus) === 'flagged';
    const isRead = message.isRead === true;
    return {
      id: stringValue(message.id),
      from: sender,
      subject: stringValue(message.subject) ?? '(no subject)',
      receivedAt: stringValue(message.receivedDateTime),
      webLink: stringValue(message.webLink),
      isRead,
      flagged,
      importance: stringValue(message.importance) ?? 'normal',
      hasAttachment: message.hasAttachments === true,
      internal: Boolean(internalDomain && emailDomain(sender) === internalDomain.toLowerCase())
    };
  });
  const focused = items
    .filter((item) => focus === 'all' || (focus === 'unread' ? !item.isRead : item.flagged))
    .sort((left, right) => {
      const score = (item: (typeof items)[number]) =>
        (item.flagged ? 8 : 0) +
        (item.importance === 'high' ? 4 : 0) +
        (!item.isRead ? 2 : 0) +
        (item.hasAttachment ? 1 : 0);
      return score(right) - score(left) || (right.receivedAt ?? '').localeCompare(left.receivedAt ?? '');
    });
  return {
    summary: {
      total: items.length,
      unread: items.filter((item) => !item.isRead).length,
      flagged: items.filter((item) => item.flagged).length,
      withAttachments: items.filter((item) => item.hasAttachment).length,
      internal: items.filter((item) => item.internal).length,
      external: items.filter((item) => !item.internal).length
    },
    focus,
    highlights: focused.slice(0, 20),
    byCategory: {
      internal_unread: focused.filter((item) => item.internal && !item.isRead),
      external_unread: focused.filter((item) => !item.internal && !item.isRead),
      flagged: focused.filter((item) => item.flagged)
    }
  };
}

function parseDateTime(value: unknown): number | undefined {
  const dateTimeValue = record(value);
  const dateTime = stringValue(dateTimeValue.dateTime) ?? stringValue(value);
  if (!dateTime) return undefined;
  const hasZone = /(?:Z|[+-]\d\d:\d\d)$/i.test(dateTime);
  const timeZone = stringValue(dateTimeValue.timeZone)?.toLowerCase();
  const zoneSuffix = timeZone === 'china standard time' ? '+08:00' : 'Z';
  const parsed = Date.parse(hasZone ? dateTime : `${dateTime}${zoneSuffix}`);
  return Number.isFinite(parsed) ? parsed : undefined;
}

const chinaUtcOffsetMs = 8 * 3_600_000;

type CalendarInterval = {
  id?: string;
  subject: string;
  start: number;
  end: number;
};

export function buildCalendarInsights(events: JsonRecord[], days: number, now = new Date()) {
  const intervals: CalendarInterval[] = events
    .flatMap((event) => {
      const start = parseDateTime(event.start);
      const end = parseDateTime(event.end);
      return start !== undefined && end !== undefined && end > start
        ? [{ id: stringValue(event.id), subject: stringValue(event.subject) ?? '(no subject)', start, end }]
        : [];
    })
    .sort((left, right) => left.start - right.start);

  const conflicts: Array<Record<string, unknown>> = [];
  for (let leftIndex = 0; leftIndex < intervals.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < intervals.length; rightIndex += 1) {
      const left = intervals[leftIndex];
      const right = intervals[rightIndex];
      if (right.start >= left.end) break;
      const overlapMinutes = Math.round((Math.min(left.end, right.end) - right.start) / 60_000);
      conflicts.push({
        event1: calendarEventSummary(left),
        event2: calendarEventSummary(right),
        overlapMinutes,
        suggestion: ''
      });
    }
  }

  const chinaNow = new Date(now.getTime() + chinaUtcOffsetMs);
  const windowStart =
    Date.UTC(chinaNow.getUTCFullYear(), chinaNow.getUTCMonth(), chinaNow.getUTCDate(), 9) - chinaUtcOffsetMs;
  const freeSlots: Array<{
    date: string;
    start: string;
    end: string;
    startDateTime: string;
    endDateTime: string;
    timeZone: string;
  }> = [];
  let busyMinutes = 0;
  for (let day = 0; day < days; day += 1) {
    const dayStart = windowStart + day * 86_400_000;
    const dayEnd = dayStart + 9 * 3_600_000;
    const busy = mergeIntervals(
      intervals
        .map((interval) => ({ start: Math.max(interval.start, dayStart), end: Math.min(interval.end, dayEnd) }))
        .filter((interval) => interval.end > interval.start)
    );
    busyMinutes += busy.reduce((sum, interval) => sum + (interval.end - interval.start) / 60_000, 0);
    let cursor = dayStart;
    for (const interval of busy) {
      if (interval.start - cursor >= 30 * 60_000) freeSlots.push(freeSlot(cursor, interval.start));
      cursor = Math.max(cursor, interval.end);
    }
    if (dayEnd - cursor >= 30 * 60_000) freeSlots.push(freeSlot(cursor, dayEnd));
  }

  const suggestedSlot = freeSlots[0];
  const conflictsWithSuggestions = conflicts.map((conflict) => ({
    ...conflict,
    suggestion: suggestedSlot
      ? `Move one event to ${suggestedSlot.date} ${suggestedSlot.start}-${suggestedSlot.end} (${suggestedSlot.timeZone}).`
      : 'No free working-hour slot of at least 30 minutes was found in the selected window.'
  }));

  return {
    totalEvents: intervals.length,
    conflicts: conflictsWithSuggestions,
    busyRatio: Number(Math.min(1, busyMinutes / (days * 9 * 60)).toFixed(2)),
    freeSlots
  };
}

function calendarEventSummary(interval: CalendarInterval) {
  return {
    id: interval.id,
    subject: interval.subject,
    start: new Date(interval.start).toISOString(),
    end: new Date(interval.end).toISOString()
  };
}

function mergeIntervals(intervals: Array<{ start: number; end: number }>) {
  const sorted = intervals.sort((left, right) => left.start - right.start);
  const merged: Array<{ start: number; end: number }> = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous || interval.start > previous.end) merged.push({ ...interval });
    else previous.end = Math.max(previous.end, interval.end);
  }
  return merged;
}

function freeSlot(start: number, end: number) {
  const chinaStart = new Date(start + chinaUtcOffsetMs).toISOString();
  const chinaEnd = new Date(end + chinaUtcOffsetMs).toISOString();
  return {
    date: chinaStart.slice(0, 10),
    start: chinaStart.slice(11, 16),
    end: chinaEnd.slice(11, 16),
    startDateTime: new Date(start).toISOString(),
    endDateTime: new Date(end).toISOString(),
    timeZone: 'China Standard Time'
  };
}

type TeamsSource = { sourceType: 'channel' | 'chat'; sourceName: string; messages: JsonRecord[] };

function stripHtml(value: string): string {
  return value
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function buildTeamsDigest(sources: TeamsSource[], limit: number) {
  const items = sources
    .flatMap((source) =>
      source.messages.map((message) => {
        const from = record(message.from);
        const sender = stringValue(record(from.user).displayName) ?? stringValue(record(from.application).displayName);
        const mentions = Array.isArray(message.mentions) ? message.mentions.length : 0;
        const attachments = Array.isArray(message.attachments) ? message.attachments.length : 0;
        return {
          sourceType: source.sourceType,
          sourceName: source.sourceName,
          messageId: stringValue(message.id),
          sender: sender ?? 'unknown',
          createdAt: stringValue(message.createdDateTime) ?? stringValue(message.lastModifiedDateTime),
          bodyPreview: stripHtml(stringValue(record(message.body).content) ?? '').slice(0, 500),
          webUrl: stringValue(message.webUrl),
          mentions,
          attachments,
          importanceScore: mentions * 3 + attachments
        };
      })
    )
    .sort(
      (left, right) =>
        right.importanceScore - left.importanceScore || (right.createdAt ?? '').localeCompare(left.createdAt ?? '')
    )
    .slice(0, limit);
  return {
    unreadStatusAvailable: false,
    statusNote:
      'Microsoft Graph channel/chat list APIs do not provide one reliable cross-source unread flag; these are recent unread candidates.',
    sourceCount: sources.length,
    candidateCount: items.length,
    items
  };
}

async function mapWithConcurrency<T, R>(
  items: T[],
  concurrency: number,
  mapper: (item: T) => Promise<R>
): Promise<R[]> {
  const results = new Array<R>(items.length);
  let next = 0;
  async function worker() {
    while (next < items.length) {
      const index = next;
      next += 1;
      results[index] = await mapper(items[index]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, () => worker()));
  return results;
}

async function smartMailDigest(config: AppConfig, hours: number, focus: MailFocus) {
  const since = new Date(Date.now() - hours * 3_600_000);
  const [messagesResponse, meResponse] = await Promise.all([listMyMessagesSince(config, since, 50), getMe(config)]);
  const me = record(meResponse);
  const internalDomain = emailDomain(stringValue(me.mail) ?? stringValue(me.userPrincipalName));
  return { hours, internalDomain, ...buildMailDigest(recordsFromCollection(messagesResponse), internalDomain, focus) };
}

async function smartCalendarConflicts(config: AppConfig, days: number) {
  return buildCalendarInsights(recordsFromCollection(await listMyCalendarEvents(config, days, 50)), days);
}

async function smartTeamsUnread(config: AppConfig, limit: number) {
  const teams = recordsFromCollection(await listMyJoinedTeams(config));
  const teamChannels = await mapWithConcurrency(teams, 4, async (team) => ({
    team,
    channels: recordsFromCollection(await listTeamChannels(config, stringValue(team.id) ?? ''))
  }));
  const channelInputs = teamChannels.flatMap(({ team, channels }) => channels.map((channel) => ({ team, channel })));
  const channelSources = await mapWithConcurrency(channelInputs, 4, async ({ team, channel }) => ({
    sourceType: 'channel' as const,
    sourceName: `${stringValue(team.displayName) ?? 'Team'} / ${stringValue(channel.displayName) ?? 'Channel'}`,
    messages: recordsFromCollection(
      await listChannelMessages(config, stringValue(team.id) ?? '', stringValue(channel.id) ?? '', Math.min(10, limit))
    )
  }));

  const chats = recordsFromCollection(await listMyChats(config, 50));
  const chatSources = await mapWithConcurrency(chats, 4, async (chat) => ({
    sourceType: 'chat' as const,
    sourceName:
      stringValue(chat.topic) ?? `${stringValue(chat.chatType) ?? 'Chat'} ${stringValue(chat.id) ?? ''}`.trim(),
    messages: recordsFromCollection(await listChatMessages(config, stringValue(chat.id) ?? '', Math.min(10, limit)))
  }));
  return buildTeamsDigest([...channelSources, ...chatSources], limit);
}

export const smartModule: ToolModule = {
  category: 'smart',
  displayName: 'Smart M365 Insights',
  description: 'Cross-item Microsoft 365 summaries and conflict insights.',
  requiredRole: 'mcp.smart',
  toolNames: ['smart_mail_digest', 'smart_calendar_conflicts', 'smart_teams_unread'],
  register(server, config) {
    server.registerTool(
      'smart_mail_digest',
      {
        title: 'Build Mail Digest',
        description: describeTool(
          'Summarize recent mail by urgency, sender domain, read state, flags, and attachments.',
          ['生成邮件摘要', '总结最近邮件', '查看重要未读邮件']
        ),
        inputSchema: {
          hours: z.number().int().min(1).max(72).default(24).describe('Number of hours to look back.'),
          focus: z.enum(['all', 'unread', 'flagged']).default('unread').describe('Messages to prioritize.')
        }
      },
      async ({ hours, focus }) => runTool(() => smartMailDigest(config, hours, focus))
    );

    server.registerTool(
      'smart_calendar_conflicts',
      {
        title: 'Find Calendar Conflicts',
        description: describeTool('Detect overlapping events and list free working-hour slots.', [
          '检查日程冲突',
          '查找空闲时间',
          '分析未来日程'
        ]),
        inputSchema: {
          days: z.number().int().min(1).max(14).default(3).describe('Number of future days to inspect.')
        }
      },
      async ({ days }) => runTool(() => smartCalendarConflicts(config, days))
    );

    server.registerTool(
      'smart_teams_unread',
      {
        title: 'Aggregate Recent Teams Messages',
        description: describeTool(
          'Aggregate recent Teams channel and chat messages as unread candidates; Graph does not expose one reliable cross-source unread flag.',
          ['聚合Teams未读消息', '总结最近Teams消息', '查看Teams消息摘要']
        ),
        inputSchema: {
          limit: z.number().int().min(5).max(50).default(20).describe('Maximum number of message candidates.')
        }
      },
      async ({ limit }) => runTool(() => smartTeamsUnread(config, limit))
    );
  }
};
