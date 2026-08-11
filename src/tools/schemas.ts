import { z } from 'zod/v4';

export const recipientSchema = z.object({
  email: z.string().email().describe('Recipient email address.'),
  name: z.string().min(1).max(120).optional().describe('Optional display name.')
});

export const eventDateTimeSchema = z.object({
  dateTime: z
    .string()
    .min(1)
    .max(80)
    .describe('Date and time, for example 2026-07-08T15:00:00. Local time without Z is recommended.'),
  timeZone: z.string().min(1).max(80).optional().describe('Time zone. Defaults to China Standard Time.')
});

export const attendeeSchema = z.object({
  email: z.string().email().describe('Attendee email address.'),
  name: z.string().min(1).max(120).optional().describe('Optional display name.'),
  type: z.enum(['required', 'optional', 'resource']).optional().describe('Attendee type. Defaults to required.')
});

export const contactEmailSchema = z.object({
  address: z.string().email().describe('Contact email address.'),
  name: z.string().min(1).max(120).optional().describe('Optional display name.')
});

export const sharePointFieldNameSchema = z
  .string()
  .min(1)
  .max(100)
  .regex(
    /^[A-Za-z_][A-Za-z0-9_]*$/,
    'Use the SharePoint internal column name, using letters, numbers, and underscores.'
  );

export const sharePointFieldsSchema = z
  .record(z.string().min(1).max(100), z.unknown())
  .refine((value) => Object.keys(value).length > 0, 'At least one SharePoint field is required.');
