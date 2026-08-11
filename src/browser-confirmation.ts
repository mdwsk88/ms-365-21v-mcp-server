import { randomBytes } from 'node:crypto';
import type { JWTPayload } from 'jose';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { operationPreviewText } from './confirmation.js';
import { errorResult } from './tools/results.js';
import type { ToolRegistryEntry } from './tools/types.js';

export type PreviewField = {
  key: string;
  label: string;
  value: string;
  multiline?: boolean;
};

export type BrowserOperationPreview = {
  kind: 'email' | 'calendar' | 'teams' | 'file' | 'delete' | 'generic';
  title: string;
  toolName: string;
  operationType: string;
  fields: PreviewField[];
  content?: PreviewField;
  approveLabel: string;
  destructive: boolean;
};

type PendingOperation = {
  userKey: string;
  operationType: string;
  expiresAt: number;
  previewText: string;
  preview: BrowserOperationPreview;
  approvalUrl: string;
  approvedAt?: number;
  browserApprovalNonce?: string;
  execute: () => Promise<CallToolResult>;
  cleanupTimer?: NodeJS.Timeout;
};

export type BrowserConfirmationView = {
  operationType: string;
  previewText: string;
  preview: BrowserOperationPreview;
  approvalUrl: string;
  expiresIn: number;
  approved: boolean;
};

export type BrowserApprovalPreparation = {
  view: BrowserConfirmationView;
  browserNonce?: string;
};

export type BrowserConfirmationRequest = {
  executed: false;
  confirmRequired: true;
  confirm_required: true;
  humanApprovalRequired: true;
  human_approval_required: true;
  confirmationStatus: 'awaiting_human_approval';
  confirmation_status: 'awaiting_human_approval';
  operationType: string;
  operation_type: string;
  preview: string;
  previewDetails: BrowserOperationPreview;
  confirmToken: string;
  confirm_token: string;
  approvalUrl: string;
  approval_url: string;
  approvalInstructions: string;
  approval_instructions: string;
  expiresIn: number;
  expires_in: number;
};

const hiddenKeys = /(?:authorization|clientsecret|password|refreshtoken|token|secret|contentbytes|filecontent)/i;

const fieldLabels: Record<string, string> = {
  subject: '主题 / Subject',
  to: '收件人 / To',
  cc: '抄送 / Cc',
  bcc: '密送 / Bcc',
  body: '正文 / Body',
  bodyIsHtml: '正文格式 / Body format',
  comment: '内容 / Comment',
  content: '内容 / Content',
  contentIsHtml: '内容格式 / Content format',
  messageId: '邮件或消息 ID / Message ID',
  eventId: '日程 ID / Event ID',
  itemId: '项目 ID / Item ID',
  driveId: '文档库 ID / Drive ID',
  teamId: '团队 ID / Team ID',
  channelId: '频道 ID / Channel ID',
  chatId: '聊天 ID / Chat ID',
  siteId: '站点 ID / Site ID',
  listId: '列表 ID / List ID',
  start: '开始时间 / Start',
  end: '结束时间 / End',
  location: '地点 / Location',
  attendees: '参与者 / Attendees',
  isOnlineMeeting: '在线会议 / Online meeting',
  type: '链接类型 / Link type',
  scope: '共享范围 / Sharing scope',
  saveToSentItems: '保存到已发送 / Save to Sent Items'
};

const operationTitles: Record<string, string> = {
  send_email: '发送邮件 / Send Mail',
  create_event: '创建日程 / Create Calendar Event',
  delete_email: '删除邮件 / Delete Mail',
  delete_event: '删除日程 / Delete Calendar Event',
  delete_drive_item: '删除 OneDrive 项目 / Delete OneDrive Item',
  delete_sharepoint_list_item: '删除 SharePoint 列表项 / Delete SharePoint List Item',
  delete_sharepoint_drive_item: '删除 SharePoint 文件 / Delete SharePoint Item',
  send_teams_message: '发送 Teams 消息 / Send Teams Message',
  create_share_link: '创建共享链接 / Create Share Link'
};

function compactText(value: string, limit = 20_000): string {
  const normalized = value.replace(/\r\n/g, '\n').trim();
  return normalized.length > limit ? `${normalized.slice(0, limit)}\n[内容已截断 / truncated]` : normalized;
}

function recipientText(value: unknown): string | undefined {
  if (!Array.isArray(value)) return undefined;
  const recipients = value.slice(0, 100).map(item => {
    if (!item || typeof item !== 'object') return compactText(String(item), 500);
    const record = item as Record<string, unknown>;
    const name = typeof record.name === 'string' ? record.name.trim() : '';
    const email = typeof record.email === 'string' ? record.email.trim() : '';
    const type = typeof record.type === 'string' ? ` (${record.type})` : '';
    return `${name && email ? `${name} <${email}>` : email || name || displayValue(record)}${type}`;
  });
  return recipients.length ? recipients.join(', ') : '无 / None';
}

function dateTimeText(value: Record<string, unknown>): string | undefined {
  if (typeof value.dateTime !== 'string') return undefined;
  const zone = typeof value.timeZone === 'string' && value.timeZone ? ` (${value.timeZone})` : '';
  return `${value.dateTime}${zone}`;
}

function displayValue(value: unknown, key = '', depth = 0): string {
  if (hiddenKeys.test(key)) return '[已隐藏 / hidden]';
  if (depth >= 5) return '[内容层级过深 / nested content]';
  if (['to', 'cc', 'bcc', 'attendees'].includes(key)) return recipientText(value) ?? '无 / None';
  if (Array.isArray(value)) {
    return value.slice(0, 100).map(item => displayValue(item, '', depth + 1)).join(', ') || '无 / None';
  }
  if (value && typeof value === 'object') {
    const record = value as Record<string, unknown>;
    const dateTime = dateTimeText(record);
    if (dateTime) return dateTime;
    return Object.entries(record)
      .slice(0, 50)
      .map(([childKey, child]) => `${fieldLabels[childKey] ?? childKey}: ${displayValue(child, childKey, depth + 1)}`)
      .join('; ');
  }
  if (typeof value === 'string') return compactText(value);
  if (typeof value === 'boolean') return value ? '是 / Yes' : '否 / No';
  if (value === null || value === undefined || value === '') return '无 / None';
  return String(value);
}

function previewKind(entry: ToolRegistryEntry): BrowserOperationPreview['kind'] {
  const operationType = entry.operationType ?? entry.name;
  if (operationType === 'send_email') return 'email';
  if (operationType.includes('event')) return 'calendar';
  if (operationType === 'send_teams_message') return 'teams';
  if (operationType.startsWith('delete_')) return 'delete';
  if (entry.category === 'drive' || entry.category === 'sharepoint') return 'file';
  return 'generic';
}

function approveLabel(operationType: string): string {
  if (operationType === 'send_email' || operationType === 'send_teams_message') return '确认并发送';
  if (operationType.startsWith('delete_')) return '确认删除';
  if (operationType === 'create_event' || operationType === 'create_share_link') return '确认并创建';
  return '确认执行';
}

function previewField(key: string, value: unknown): PreviewField {
  let rendered: string;
  if (key === 'bodyIsHtml' || key === 'contentIsHtml') {
    rendered = value === true ? 'HTML' : '纯文本 / Plain text';
  } else {
    rendered = displayValue(value, key);
  }
  return {
    key,
    label: fieldLabels[key] ?? key,
    value: rendered,
    multiline: false
  };
}

export function createBrowserOperationPreview(
  entry: ToolRegistryEntry,
  parameters: Record<string, unknown>
): BrowserOperationPreview {
  const operationType = entry.operationType ?? entry.name;
  const contentKey = ['body', 'comment', 'content'].find(key => parameters[key] !== undefined);
  const fields =
    entry.name === 'mail_send'
      ? [
          previewField('to', parameters.to),
          previewField('cc', parameters.cc),
          previewField('bcc', parameters.bcc),
          previewField('subject', parameters.subject),
          previewField('bodyIsHtml', parameters.bodyIsHtml ?? false),
          previewField('saveToSentItems', parameters.saveToSentItems ?? true)
        ]
      : Object.entries(parameters)
          .filter(([key]) => key !== contentKey)
          .slice(0, 60)
          .map(([key, value]) => previewField(key, value));
  const content = contentKey
    ? {
        key: contentKey,
        label: fieldLabels[contentKey] ?? contentKey,
        value: displayValue(parameters[contentKey], contentKey),
        multiline: true
      }
    : undefined;

  return {
    kind: previewKind(entry),
    title: operationTitles[operationType] ?? entry.title ?? entry.name,
    toolName: entry.name,
    operationType,
    fields,
    content,
    approveLabel: approveLabel(operationType),
    destructive: operationType.startsWith('delete_')
  };
}

export function confirmationUserKey(claims: JWTPayload | undefined): string {
  const oid = typeof claims?.oid === 'string' ? claims.oid : undefined;
  return oid ?? claims?.sub ?? 'local-or-anonymous';
}

function pendingText(request: BrowserConfirmationRequest, fallbackDetail?: string): string {
  const detail = fallbackDetail ? '\n客户端确认能力不可用，已自动切换到网页确认。' : '';
  return [
    `操作尚未执行。当前客户端不能完成原生确认，已生成安全网页预览。${detail}`,
    '请把下面的链接展示给用户并停止，不要改用草稿或其他写工具绕过确认：',
    request.approvalUrl,
    '',
    '用户在网页批准后，再调用 confirm_execute，并传入：',
    `confirmToken: ${request.confirmToken}`,
    '',
    request.preview
  ].join('\n');
}

export function browserConfirmationResult(
  request: BrowserConfirmationRequest,
  fallbackDetail?: string
): CallToolResult {
  return {
    content: [{ type: 'text', text: pendingText(request, fallbackDetail) }],
    structuredContent: request
  };
}

export class BrowserConfirmationStore {
  private readonly pending = new Map<string, PendingOperation>();

  constructor(private readonly clock: () => number = Date.now) {}

  create(
    entry: ToolRegistryEntry,
    parameters: Record<string, unknown>,
    userKey: string,
    ttlSeconds: number,
    execute: () => Promise<CallToolResult>,
    approvalBaseUrl = 'http://127.0.0.1:3000'
  ): BrowserConfirmationRequest {
    this.pruneExpired();
    const confirmToken = `confirm_${randomBytes(32).toString('base64url')}`;
    const operationType = entry.operationType ?? entry.name;
    const preview = createBrowserOperationPreview(entry, parameters);
    const previewText = operationPreviewText(entry, parameters);
    const approvalUrl = `${approvalBaseUrl.replace(/\/+$/, '')}/confirm/${encodeURIComponent(confirmToken)}`;
    const operation: PendingOperation = {
      userKey,
      operationType,
      expiresAt: this.clock() + ttlSeconds * 1000,
      previewText,
      preview,
      approvalUrl,
      execute
    };
    operation.cleanupTimer = setTimeout(() => {
      if (this.pending.get(confirmToken) === operation) this.delete(confirmToken);
    }, ttlSeconds * 1000);
    operation.cleanupTimer.unref?.();
    this.pending.set(confirmToken, operation);

    const instructions =
      'Show approval_url to the user and stop. Do not use a draft or another write tool. After the user approves the page, call confirm_execute with confirm_token.';
    return {
      executed: false,
      confirmRequired: true,
      confirm_required: true,
      humanApprovalRequired: true,
      human_approval_required: true,
      confirmationStatus: 'awaiting_human_approval',
      confirmation_status: 'awaiting_human_approval',
      operationType,
      operation_type: operationType,
      preview: previewText,
      previewDetails: preview,
      confirmToken,
      confirm_token: confirmToken,
      approvalUrl,
      approval_url: approvalUrl,
      approvalInstructions: instructions,
      approval_instructions: instructions,
      expiresIn: ttlSeconds,
      expires_in: ttlSeconds
    };
  }

  prepareBrowserApproval(confirmToken: string): BrowserApprovalPreparation | undefined {
    const operation = this.activeOperation(confirmToken);
    if (!operation) return undefined;
    if (operation.approvedAt !== undefined) return { view: this.approvalView(operation) };

    operation.browserApprovalNonce = randomBytes(32).toString('base64url');
    return {
      view: this.approvalView(operation),
      browserNonce: operation.browserApprovalNonce
    };
  }

  approveFromBrowser(confirmToken: string, browserNonce: string): BrowserConfirmationView | undefined {
    const operation = this.activeOperation(confirmToken);
    if (!operation || operation.approvedAt !== undefined || operation.browserApprovalNonce !== browserNonce) {
      return undefined;
    }
    operation.approvedAt = this.clock();
    operation.browserApprovalNonce = undefined;
    return this.approvalView(operation);
  }

  rejectFromBrowser(confirmToken: string, browserNonce: string): BrowserConfirmationView | undefined {
    const operation = this.activeOperation(confirmToken);
    if (!operation || operation.approvedAt !== undefined || operation.browserApprovalNonce !== browserNonce) {
      return undefined;
    }
    const view = this.approvalView(operation);
    this.delete(confirmToken);
    return view;
  }

  async execute(confirmToken: string, userKey: string): Promise<CallToolResult> {
    const operation = this.pending.get(confirmToken);
    if (!operation) {
      return errorResult('invalid_confirm_token', 'Confirmation token is unknown, expired, rejected, or already used.');
    }
    if (operation.expiresAt <= this.clock()) {
      this.delete(confirmToken);
      return errorResult('expired_confirm_token', 'Confirmation token has expired.');
    }
    if (operation.userKey !== userKey) {
      return errorResult('confirm_user_mismatch', 'Confirmation token belongs to a different signed-in user.');
    }
    if (operation.approvedAt === undefined) return this.humanApprovalRequiredResult(operation);

    this.delete(confirmToken);
    return operation.execute();
  }

  size(): number {
    this.pruneExpired();
    return this.pending.size;
  }

  clear(): void {
    for (const operation of this.pending.values()) {
      if (operation.cleanupTimer) clearTimeout(operation.cleanupTimer);
    }
    this.pending.clear();
  }

  private pruneExpired(): void {
    const now = this.clock();
    for (const [token, operation] of this.pending) {
      if (operation.expiresAt <= now) this.delete(token);
    }
  }

  private activeOperation(confirmToken: string): PendingOperation | undefined {
    const operation = this.pending.get(confirmToken);
    if (!operation) return undefined;
    if (operation.expiresAt <= this.clock()) {
      this.delete(confirmToken);
      return undefined;
    }
    return operation;
  }

  private approvalView(operation: PendingOperation): BrowserConfirmationView {
    return {
      operationType: operation.operationType,
      previewText: operation.previewText,
      preview: operation.preview,
      approvalUrl: operation.approvalUrl,
      expiresIn: Math.max(0, Math.ceil((operation.expiresAt - this.clock()) / 1000)),
      approved: operation.approvedAt !== undefined
    };
  }

  private humanApprovalRequiredResult(operation: PendingOperation): CallToolResult {
    const output = {
      error: {
        code: 'human_confirmation_required',
        message: 'A human must approve the operation in the confirmation page before execution.'
      },
      executed: false,
      confirmationStatus: 'awaiting_human_approval',
      confirmation_status: 'awaiting_human_approval',
      approvalUrl: operation.approvalUrl,
      approval_url: operation.approvalUrl,
      preview: operation.previewText,
      previewDetails: operation.preview
    };
    return {
      content: [
        {
          type: 'text',
          text: `网页确认尚未完成，操作未执行。请让用户先打开并批准：\n${operation.approvalUrl}`
        }
      ],
      structuredContent: output,
      isError: true
    };
  }

  private delete(token: string): void {
    const operation = this.pending.get(token);
    if (operation?.cleanupTimer) clearTimeout(operation.cleanupTimer);
    this.pending.delete(token);
  }
}

export const browserConfirmationStore = new BrowserConfirmationStore();
