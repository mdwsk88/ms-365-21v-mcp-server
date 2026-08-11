import fs from 'node:fs/promises';
import path from 'node:path';
import type { AppConfig } from './config.js';

export interface AuditEntry {
  timestamp: string;
  requestId: string;
  userId: string;
  userDisplayName?: string;
  toolName: string;
  toolCategory: string;
  isWriteOperation: boolean;
  parameters: Record<string, unknown>;
  duration: number;
  success: boolean;
  errorCode?: string;
  graphScopes?: string[];
  eventType?: string;
}

type Aggregate = { calls: number; errors: number; durationMs: number };

export interface AuditSummary {
  generatedAt: string;
  days: number;
  totalCalls: number;
  successfulCalls: number;
  failedCalls: number;
  byDay: Record<string, Aggregate>;
  byUser: Record<string, Aggregate>;
  byTool: Record<string, Aggregate>;
}

const sensitiveKeys = new Set([
  'authorization',
  'body',
  'clientsecret',
  'comment',
  'content',
  'contentbytes',
  'filecontent',
  'password',
  'refreshtoken',
  'token'
]);

function isSensitiveKey(key: string): boolean {
  const normalized = key.toLowerCase().replace(/[^a-z0-9]/g, '');
  return sensitiveKeys.has(normalized) || normalized.endsWith('token') || normalized.endsWith('secret');
}

function redactValue(value: unknown, key?: string, depth = 0): unknown {
  if (key && isSensitiveKey(key)) return '[REDACTED]';
  if (depth >= 8) return '[TRUNCATED]';
  if (Array.isArray(value)) return value.map((item) => redactValue(item, undefined, depth + 1));
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>).map(([childKey, childValue]) => [
        childKey,
        redactValue(childValue, childKey, depth + 1)
      ])
    );
  }
  if (typeof value === 'string' && value.length > 1000) return `${value.slice(0, 1000)}[TRUNCATED]`;
  return value;
}

export function redactAuditParameters(parameters: Record<string, unknown>): Record<string, unknown> {
  return redactValue(parameters) as Record<string, unknown>;
}

function storageLocation(configuredPath: string): { directory: string; prefix: string } {
  if (path.extname(configuredPath).toLowerCase() === '.log') {
    return { directory: path.dirname(configuredPath), prefix: path.basename(configuredPath, '.log') };
  }
  return { directory: configuredPath, prefix: 'audit' };
}

function dateKey(date = new Date()): string {
  return date.toISOString().slice(0, 10);
}

function aggregate(target: Record<string, Aggregate>, key: string, entry: AuditEntry): void {
  const value = target[key] ?? { calls: 0, errors: 0, durationMs: 0 };
  value.calls += 1;
  value.errors += entry.success ? 0 : 1;
  value.durationMs += entry.duration;
  target[key] = value;
}

export class AuditLogger {
  private writeChain: Promise<void> = Promise.resolve();
  private readonly directory: string;
  private readonly prefix: string;

  constructor(
    readonly enabled: boolean,
    configuredPath: string
  ) {
    const location = storageLocation(configuredPath);
    this.directory = location.directory;
    this.prefix = location.prefix;
  }

  filePath(date = new Date()): string {
    return path.join(this.directory, `${this.prefix}-${dateKey(date)}.log`);
  }

  async write(entry: AuditEntry): Promise<void> {
    if (!this.enabled) return;
    const sanitized: AuditEntry = {
      ...entry,
      parameters: redactAuditParameters(entry.parameters),
      graphScopes: entry.graphScopes ? [...entry.graphScopes] : undefined
    };
    const line = `${JSON.stringify(sanitized)}\n`;
    const pendingWrite = this.writeChain
      .catch(() => undefined)
      .then(async () => {
        await fs.mkdir(this.directory, { recursive: true, mode: 0o700 });
        await fs.chmod(this.directory, 0o700);
        const handle = await fs.open(this.filePath(new Date(entry.timestamp)), 'a', 0o600);
        try {
          await handle.write(line);
          await handle.chmod(0o600);
        } finally {
          await handle.close();
        }
      });
    this.writeChain = pendingWrite;
    await pendingWrite;
  }

  async summary(days = 7, now = new Date()): Promise<AuditSummary> {
    const summary: AuditSummary = {
      generatedAt: now.toISOString(),
      days,
      totalCalls: 0,
      successfulCalls: 0,
      failedCalls: 0,
      byDay: {},
      byUser: {},
      byTool: {}
    };
    if (!this.enabled) return summary;

    for (let offset = 0; offset < days; offset += 1) {
      const date = new Date(now);
      date.setUTCDate(date.getUTCDate() - offset);
      let contents: string;
      try {
        contents = await fs.readFile(this.filePath(date), 'utf8');
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') continue;
        throw error;
      }
      for (const line of contents.split('\n')) {
        if (!line.trim()) continue;
        let entry: AuditEntry;
        try {
          entry = JSON.parse(line) as AuditEntry;
        } catch {
          continue;
        }
        if (entry.eventType) continue;
        summary.totalCalls += 1;
        summary.successfulCalls += entry.success ? 1 : 0;
        summary.failedCalls += entry.success ? 0 : 1;
        aggregate(summary.byDay, entry.timestamp.slice(0, 10), entry);
        aggregate(summary.byUser, entry.userId || 'unknown', entry);
        aggregate(summary.byTool, entry.toolName, entry);
      }
    }
    return summary;
  }
}

const loggers = new Map<string, AuditLogger>();

export function getAuditLogger(config: AppConfig): AuditLogger {
  const key = `${config.auditLogEnabled}:${config.auditLogPath}`;
  let logger = loggers.get(key);
  if (!logger) {
    logger = new AuditLogger(config.auditLogEnabled, config.auditLogPath);
    loggers.set(key, logger);
  }
  return logger;
}
