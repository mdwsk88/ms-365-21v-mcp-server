import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
  ElicitResultSchema,
  type CallToolResult,
  type ElicitRequestFormParams,
  type ElicitResult
} from '@modelcontextprotocol/sdk/types.js';
import type { ToolRegistryEntry } from './tools/types.js';
import type { AppConfig } from './config.js';

type RequestHandlerExtraLike = {
  sendRequest?: (request: unknown, resultSchema: unknown, options?: unknown) => Promise<unknown>;
};

export type FormElicitor = {
  isSupported(): boolean;
  request(params: ElicitRequestFormParams): Promise<ElicitResult>;
};

export type OperationConfirmation =
  | { status: 'approved' }
  | { status: 'stopped'; result: CallToolResult }
  | { status: 'browser_fallback'; detail?: string };

const hiddenKeys = /(?:authorization|clientsecret|password|refreshtoken|token|secret|contentbytes|filecontent)/i;

const labels: Record<string, string> = {
  subject: '主题 / Subject',
  to: '收件人 / To',
  cc: '抄送 / Cc',
  bcc: '密送 / Bcc',
  body: '正文 / Body',
  comment: '内容 / Comment',
  content: '内容 / Content',
  messageId: '邮件或消息 ID / Message ID',
  eventId: '日程 ID / Event ID',
  itemId: '项目 ID / Item ID',
  teamId: '团队 ID / Team ID',
  channelId: '频道 ID / Channel ID',
  siteId: '站点 ID / Site ID',
  listId: '列表 ID / List ID',
  start: '开始时间 / Start',
  end: '结束时间 / End',
  location: '地点 / Location',
  attendees: '参与者 / Attendees',
  type: '类型 / Type',
  scope: '范围 / Scope',
  saveToSentItems: '保存到已发送 / Save to Sent Items'
};

function compactText(value: string, limit = 1200): string {
  const normalized = value.replace(/\r\n/g, '\n').trim();
  return normalized.length > limit ? `${normalized.slice(0, limit)}\n[内容已截断 / truncated]` : normalized;
}

function recipientText(value: unknown): string | undefined {
  if (!Array.isArray(value)) return undefined;
  const recipients = value.slice(0, 50).map(item => {
    if (!item || typeof item !== 'object') return String(item);
    const record = item as Record<string, unknown>;
    const name = typeof record.name === 'string' ? record.name.trim() : '';
    const email = typeof record.email === 'string' ? record.email.trim() : '';
    return name && email ? `${name} <${email}>` : email || name || objectText(record);
  });
  return recipients.length ? recipients.join(', ') : '无 / None';
}

function objectText(value: Record<string, unknown>): string {
  return Object.entries(value)
    .slice(0, 20)
    .map(([key, child]) => `${labels[key] ?? key}: ${valueText(child, key)}`)
    .join('; ');
}

function valueText(value: unknown, key = ''): string {
  if (hiddenKeys.test(key)) return '[已隐藏 / hidden]';
  const recipients = recipientText(value);
  if (recipients !== undefined) return recipients;
  if (Array.isArray(value)) return value.slice(0, 50).map(item => valueText(item)).join(', ') || '无 / None';
  if (value && typeof value === 'object') return objectText(value as Record<string, unknown>);
  if (typeof value === 'string') return compactText(value);
  if (typeof value === 'boolean') return value ? '是 / Yes' : '否 / No';
  if (value === null || value === undefined || value === '') return '无 / None';
  return String(value);
}

export function operationPreviewText(
  entry: ToolRegistryEntry,
  parameters: Record<string, unknown>
): string {
  const parameterLines = Object.entries(parameters)
    .slice(0, 40)
    .map(([key, value]) => {
      const rendered = valueText(value, key);
      return ['body', 'comment', 'content'].includes(key)
        ? `${labels[key] ?? key}:\n${rendered}`
        : `${labels[key] ?? key}: ${rendered}`;
    });

  return [
    `操作 / Operation: ${entry.title ?? entry.name}`,
    `工具 / Tool: ${entry.name}`,
    `操作类型 / Operation type: ${entry.operationType ?? entry.name}`,
    ...parameterLines
  ].join('\n');
}

export function isConfirmationOperationEnabled(confirmOperations: string[], operationType: string): boolean {
  const normalized = confirmOperations.map(value => value.trim().toLowerCase()).filter(Boolean);
  return normalized.includes('*') || normalized.includes('all') || normalized.includes(operationType.toLowerCase());
}

const sendOperationTypes = new Set(['send_email', 'send_teams_message']);

export function operationRequiresConfirmation(
  config: Pick<AppConfig, 'confirmOperations' | 'sendMode'>,
  operationType: string
): boolean {
  const normalized = operationType.trim().toLowerCase();
  if (sendOperationTypes.has(normalized)) {
    return config.sendMode === 'confirm';
  }
  return isConfirmationOperationEnabled(config.confirmOperations, normalized);
}

export function hasConfirmationPolicy(
  config: Pick<AppConfig, 'confirmOperations' | 'sendMode'>
): boolean {
  if (config.sendMode === 'confirm') return true;
  return config.confirmOperations.some(value => {
    const normalized = value.trim().toLowerCase();
    return Boolean(normalized) && !sendOperationTypes.has(normalized);
  });
}

export function createFormElicitor(
  server: McpServer,
  extra: RequestHandlerExtraLike = {}
): FormElicitor {
  return {
    isSupported: () => {
      const capability = server.server.getClientCapabilities()?.elicitation;
      if (!capability) return false;
      const modes = capability as Record<string, unknown>;
      return Boolean(modes.form) || (!('form' in modes) && !('url' in modes));
    },
    async request(params) {
      if (extra.sendRequest) {
        return (await extra.sendRequest(
          { method: 'elicitation/create', params },
          ElicitResultSchema
        )) as ElicitResult;
      }
      return server.server.request({ method: 'elicitation/create', params }, ElicitResultSchema);
    }
  };
}

function stoppedResult(
  status: 'declined' | 'cancelled' | 'not_confirmed',
  preview: string
): CallToolResult {
  const message =
    status === 'declined'
      ? '用户已拒绝该操作，未执行任何写入。 / The user declined. No write was performed.'
      : status === 'cancelled'
        ? '用户已取消确认，未执行任何写入。 / Confirmation was cancelled. No write was performed.'
        : '用户没有勾选“确认执行”，未执行任何写入。 / Confirmation was not selected. No write was performed.';
  return {
    content: [{ type: 'text', text: `${message}\n\n${preview}` }],
    structuredContent: {
      executed: false,
      confirmationStatus: status,
      confirmation_status: status,
      preview
    }
  };
}

export async function requestOperationConfirmation(
  elicitor: FormElicitor,
  entry: ToolRegistryEntry,
  parameters: Record<string, unknown>
): Promise<OperationConfirmation> {
  const preview = operationPreviewText(entry, parameters);
  if (!elicitor.isSupported()) {
    return { status: 'browser_fallback' };
  }

  let response: ElicitResult;
  try {
    response = await elicitor.request({
      mode: 'form',
      message: [
        '请确认以下敏感操作 / Please review this sensitive operation',
        '',
        preview,
        '',
        '勾选“确认执行”并接受后，服务器会立即以当前登录用户身份执行。 / Check Confirm and accept to execute now as the signed-in user.'
      ].join('\n'),
      requestedSchema: {
        type: 'object',
        properties: {
          confirm: {
            type: 'boolean',
            title: '确认执行 / Confirm execution',
            description: '我已检查上述目标和内容，并同意立即执行。 / I reviewed the target and content and approve immediate execution.',
            default: false
          }
        },
        required: ['confirm']
      }
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return { status: 'browser_fallback', detail: message };
  }

  if (response.action === 'decline') {
    return { status: 'stopped', result: stoppedResult('declined', preview) };
  }
  if (response.action === 'cancel') {
    return { status: 'stopped', result: stoppedResult('cancelled', preview) };
  }
  if (response.content?.confirm !== true) {
    return { status: 'stopped', result: stoppedResult('not_confirmed', preview) };
  }
  return { status: 'approved' };
}
