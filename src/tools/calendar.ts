import { z } from 'zod/v4';
import {
  addEventFileAttachment,
  cancelCalendarEvent,
  createCalendar,
  deleteCalendar,
  deleteEventAttachment,
  getSchedule,
  listCalendarEventInstances,
  listCalendarView,
  listEventAttachments,
  updateCalendar
} from '../graph-extended.js';
import {
  createCalendarEvent,
  deleteCalendarEvent,
  getCalendarEvent,
  listMyCalendars,
  listMyCalendarEvents,
  respondToCalendarEvent,
  updateCalendarEvent
} from '../graph.js';
import { attendeeSchema, eventDateTimeSchema } from './schemas.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

export const calendarModule: ToolModule = {
  category: 'calendar',
  displayName: 'Calendar',
  description: 'Calendar read, scheduling, and meeting response tools.',
  requiredRole: 'mcp.calendar',
  toolNames: [
    'calendar_list_events',
    'calendar_list_view',
    'calendar_list_calendars',
    'calendar_create_calendar',
    'calendar_update_calendar',
    'calendar_delete_calendar',
    'calendar_get_event',
    'calendar_list_event_instances',
    'calendar_get_schedule',
    'calendar_create_event',
    'calendar_update_event',
    'calendar_cancel_event',
    'calendar_delete_event',
    'calendar_respond_to_event',
    'calendar_list_attachments',
    'calendar_add_file_attachment',
    'calendar_delete_attachment'
  ],
  register(server, config) {
    server.registerTool(
      'calendar_list_events',
      {
        title: 'List My Calendar Events',
        description: describeTool(
          'List upcoming calendar events for the signed-in user. Use for today, this week, or upcoming meeting queries. Requires delegated Calendars.Read.',
          ['查看我的日历', '查看今天会议', '查看本周日程', '查看接下来的会议']
        ),
        inputSchema: {
          daysAhead: z.number().int().min(1).max(30).optional().describe('How many days ahead to query, from 1 to 30.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of events to return, from 1 to 50.')
        }
      },
      async ({ daysAhead, top }) => runTool(async () => listMyCalendarEvents(config, daysAhead, top))
    );

    server.registerTool(
      'calendar_list_view',
      {
        title: 'List Calendar View',
        description: describeTool(
          'List single and recurring event instances in an explicit date-time range. Requires delegated Calendars.Read.',
          ['按时间范围查看日历', '查看指定日期的会议', '展开重复日程']
        ),
        inputSchema: {
          startDateTime: z.string().min(1).max(80).describe('Range start in ISO 8601 format.'),
          endDateTime: z.string().min(1).max(80).describe('Range end in ISO 8601 format.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of events to return.'),
          timeZone: z.string().min(1).max(80).optional().describe('Response time zone. Defaults to China Standard Time.')
        }
      },
      async ({ startDateTime, endDateTime, top, timeZone }) =>
        runTool(async () => listCalendarView(config, startDateTime, endDateTime, top, timeZone))
    );

    server.registerTool(
      'calendar_list_calendars',
      {
        title: 'List My Calendars',
        description: describeTool('List calendars for the signed-in user. Requires delegated Calendars.Read.', [
          '查看日历列表',
          '列出我的日历',
          '有哪些日历'
        ]),
        inputSchema: {
          top: z.number().int().min(1).max(100).optional().describe('Number of calendars to return, from 1 to 100.')
        }
      },
      async ({ top }) => runTool(async () => listMyCalendars(config, top))
    );

    server.registerTool(
      'calendar_create_calendar',
      {
        title: 'Create Calendar',
        description: describeTool('Create an additional calendar for the signed-in user.', [
          '创建新日历',
          '新建项目日历',
          '添加一个日历'
        ]),
        inputSchema: {
          name: z.string().min(1).max(255).describe('Calendar name.')
        }
      },
      async ({ name }) => runTool(async () => createCalendar(config, name))
    );

    server.registerTool(
      'calendar_update_calendar',
      {
        title: 'Rename Calendar',
        description: describeTool('Rename one of the signed-in user calendars.', ['重命名日历', '修改日历名称']),
        inputSchema: {
          calendarId: z.string().min(1).max(300).describe('Calendar ID.'),
          name: z.string().min(1).max(255).describe('New calendar name.')
        }
      },
      async ({ calendarId, name }) => runTool(async () => updateCalendar(config, calendarId, name))
    );

    server.registerTool(
      'calendar_delete_calendar',
      {
        title: 'Delete Calendar',
        description: describeTool('Delete an additional calendar and its events.', ['删除日历', '移除项目日历']),
        inputSchema: {
          calendarId: z.string().min(1).max(300).describe('Calendar ID to delete.')
        }
      },
      async ({ calendarId }) => runTool(async () => deleteCalendar(config, calendarId))
    );

    server.registerTool(
      'calendar_get_event',
      {
        title: 'Get Calendar Event',
        description: describeTool('Read details for a calendar event. Requires delegated Calendars.Read.', [
          '查看会议详情',
          '查看日程详情',
          '读取日历事件'
        ]),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Calendar event ID.')
        }
      },
      async ({ eventId }) => runTool(async () => getCalendarEvent(config, eventId))
    );

    server.registerTool(
      'calendar_list_event_instances',
      {
        title: 'List Recurring Event Instances',
        description: describeTool('List occurrences of a recurring event in a date-time range.', [
          '查看重复会议实例',
          '查看周期日程的每次会议',
          '展开循环会议'
        ]),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Recurring series master event ID.'),
          startDateTime: z.string().min(1).max(80).describe('Range start in ISO 8601 format.'),
          endDateTime: z.string().min(1).max(80).describe('Range end in ISO 8601 format.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of instances to return.')
        }
      },
      async ({ eventId, startDateTime, endDateTime, top }) =>
        runTool(async () => listCalendarEventInstances(config, eventId, startDateTime, endDateTime, top))
    );

    server.registerTool(
      'calendar_get_schedule',
      {
        title: 'Get Free Busy Schedule',
        description: describeTool(
          'Get free/busy availability for users, rooms, or distribution lists in a time range. Supported in 21V and requires delegated Calendars.Read.',
          ['查看同事忙闲', '查询会议室是否空闲', '找共同空闲时间']
        ),
        inputSchema: {
          schedules: z.array(z.string().email()).min(1).max(20).describe('User, room, or list SMTP addresses.'),
          start: eventDateTimeSchema.describe('Availability range start.'),
          end: eventDateTimeSchema.describe('Availability range end.'),
          availabilityViewInterval: z
            .number()
            .int()
            .min(5)
            .max(1440)
            .optional()
            .describe('Availability slot size in minutes. Defaults to 30.')
        }
      },
      async ({ schedules, start, end, availabilityViewInterval }) =>
        runTool(async () => getSchedule(config, schedules, start, end, availabilityViewInterval))
    );

    server.registerTool(
      'calendar_create_event',
      {
        title: 'Create Calendar Event',
        description: describeTool(
          'Create an event or meeting in the signed-in user default calendar. Requires delegated Calendars.ReadWrite.',
          ['创建会议', '新建日程', '安排会议', '创建日历事件']
        ),
        inputSchema: {
          subject: z.string().min(1).max(255).describe('Event subject.'),
          start: eventDateTimeSchema.describe('Start time.'),
          end: eventDateTimeSchema.describe('End time.'),
          body: z.string().max(20000).optional().describe('Event body. May contain HTML.'),
          location: z.string().max(500).optional().describe('Location.'),
          attendees: z.array(attendeeSchema).max(100).optional().describe('Attendees.'),
          isOnlineMeeting: z.boolean().optional().describe('Whether to create an online meeting.')
        }
      },
      async (input) => runTool(async () => createCalendarEvent(config, input))
    );

    server.registerTool(
      'calendar_update_event',
      {
        title: 'Update Calendar Event',
        description: describeTool(
          'Update a calendar event. Provide only the fields to change. Requires delegated Calendars.ReadWrite.',
          ['修改会议', '更新日程', '改会议时间', '调整日历事件']
        ),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Calendar event ID.'),
          subject: z.string().min(1).max(255).optional().describe('New subject.'),
          start: eventDateTimeSchema.optional().describe('New start time.'),
          end: eventDateTimeSchema.optional().describe('New end time.'),
          body: z.string().max(20000).optional().describe('New body. May contain HTML.'),
          location: z.string().max(500).optional().describe('New location.'),
          attendees: z.array(attendeeSchema).max(100).optional().describe('New attendee list.'),
          isOnlineMeeting: z.boolean().optional().describe('Whether this should be an online meeting.')
        }
      },
      async ({ eventId, ...input }) => runTool(async () => updateCalendarEvent(config, eventId, input))
    );

    server.registerTool(
      'calendar_cancel_event',
      {
        title: 'Cancel Calendar Event',
        description: describeTool(
          'Cancel a meeting organized by the signed-in user and notify attendees. Requires delegated Calendars.ReadWrite.',
          ['取消我组织的会议', '取消会议并通知参会人', '发送会议取消通知']
        ),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Event ID.'),
          comment: z.string().max(10000).optional().describe('Optional cancellation message.')
        }
      },
      async ({ eventId, comment }) => runTool(async () => cancelCalendarEvent(config, eventId, comment))
    );

    server.registerTool(
      'calendar_delete_event',
      {
        title: 'Delete Calendar Event',
        description: describeTool('Delete a calendar event. Requires delegated Calendars.ReadWrite.', [
          '删除会议',
          '取消日程',
          '删除日历事件'
        ]),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Calendar event ID.')
        }
      },
      async ({ eventId }) => runTool(async () => deleteCalendarEvent(config, eventId))
    );

    server.registerTool(
      'calendar_respond_to_event',
      {
        title: 'Respond To Calendar Event',
        description: describeTool(
          'Accept, tentatively accept, or decline a meeting invitation. Requires delegated Calendars.ReadWrite.',
          ['接受会议邀请', '拒绝会议邀请', '暂定会议', '回复会议邀请']
        ),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Calendar event ID.'),
          response: z.enum(['accept', 'tentativelyAccept', 'decline']).describe('Response action.'),
          comment: z.string().max(10000).optional().describe('Optional response comment.'),
          sendResponse: z
            .boolean()
            .optional()
            .describe('Whether to send the response to the organizer. Defaults to true.')
        }
      },
      async ({ eventId, response, comment, sendResponse }) =>
        runTool(async () => respondToCalendarEvent(config, eventId, response, comment, sendResponse))
    );

    server.registerTool(
      'calendar_list_attachments',
      {
        title: 'List Calendar Event Attachments',
        description: describeTool('List attachment metadata for a calendar event.', [
          '查看会议附件',
          '列出日程附件',
          '会议有没有附件'
        ]),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Event ID.')
        }
      },
      async ({ eventId }) => runTool(async () => listEventAttachments(config, eventId))
    );

    server.registerTool(
      'calendar_add_file_attachment',
      {
        title: 'Add File Attachment To Event',
        description: describeTool('Add a base64-encoded file under 3 MB to a calendar event.', [
          '给会议添加附件',
          '给日程附加文件'
        ]),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Event ID.'),
          name: z.string().min(1).max(255).describe('Attachment file name.'),
          contentType: z.string().min(1).max(200).describe('MIME type.'),
          contentBase64: z.string().min(1).max(4_200_000).describe('Base64 content, under 3 MB decoded.'),
          isInline: z.boolean().optional().describe('Whether this is an inline attachment.'),
          contentId: z.string().min(1).max(255).optional().describe('Optional inline content ID.')
        }
      },
      async ({ eventId, name, contentType, contentBase64, isInline, contentId }) =>
        runTool(async () =>
          addEventFileAttachment(config, eventId, name, contentType, contentBase64, isInline, contentId)
        )
    );

    server.registerTool(
      'calendar_delete_attachment',
      {
        title: 'Delete Calendar Event Attachment',
        description: describeTool('Delete an attachment from a calendar event.', ['删除会议附件', '移除日程附件']),
        inputSchema: {
          eventId: z.string().min(1).max(300).describe('Event ID.'),
          attachmentId: z.string().min(1).max(500).describe('Attachment ID.')
        }
      },
      async ({ eventId, attachmentId }) =>
        runTool(async () => deleteEventAttachment(config, eventId, attachmentId))
    );
  }
};
