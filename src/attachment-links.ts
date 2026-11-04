import { randomBytes } from 'node:crypto';
import { Readable } from 'node:stream';
import type { ReadableStream as WebReadableStream } from 'node:stream/web';
import type { AppConfig } from './config.js';
import { mountedRoutePaths } from './http-routes.js';

/**
 * Single-use download links for Graph byte endpoints that have no pre-authenticated URL of their own,
 * such as mail attachments. A link carries a random ticket; the signed-in user's Graph token stays in
 * server memory with it and is used once, when the client fetches the link, then discarded.
 */

type AttachmentLink = {
  graphPath: string;
  accessToken: string;
  name: string;
  contentType: string;
  userKey: string;
  expiresAt: number;
};

const maxLinksPerUser = 20;
const maxLinks = 500;

export class AttachmentLinkStore {
  private readonly links = new Map<string, AttachmentLink>();

  constructor(private readonly now: () => number = () => Date.now()) {}

  create(link: Omit<AttachmentLink, 'expiresAt'>, ttlSeconds: number): { ticket: string; expiresAt: number } {
    this.sweep();
    const ticket = randomBytes(32).toString('base64url');
    const expiresAt = this.now() + ttlSeconds * 1000;
    const own = [...this.links].filter(([, existing]) => existing.userKey === link.userKey);
    for (const [oldTicket] of own.slice(0, Math.max(0, own.length - maxLinksPerUser + 1))) this.links.delete(oldTicket);
    while (this.links.size >= maxLinks) this.links.delete(this.links.keys().next().value as string);
    this.links.set(ticket, { ...link, expiresAt });
    return { ticket, expiresAt };
  }

  /** Returns the link once and forgets it; expired or unknown tickets return undefined. */
  redeem(ticket: string): AttachmentLink | undefined {
    const link = this.links.get(ticket);
    if (!link) return undefined;
    this.links.delete(ticket);
    return link.expiresAt > this.now() ? link : undefined;
  }

  get size(): number {
    return this.links.size;
  }

  clear(): void {
    this.links.clear();
  }

  private sweep(): void {
    const now = this.now();
    for (const [ticket, link] of this.links) if (link.expiresAt <= now) this.links.delete(ticket);
  }
}

export const attachmentLinkStore = new AttachmentLinkStore();

export function attachmentLinkUrl(config: AppConfig, ticket: string): string {
  const base = (config.publicBaseUrl ?? config.oauthBridgeIssuer).replace(/\/+$/, '');
  return `${base}/attachments/${encodeURIComponent(ticket)}`;
}

/** Links need a reachable HTTP route, which stdio deployments do not have. */
export function attachmentLinksAvailable(config: AppConfig): boolean {
  return config.transport === 'http';
}

function attachmentLog(event: string, details: Record<string, unknown> = {}): void {
  console.error(JSON.stringify({ component: 'attachment_link', event, ...details }));
}

function contentDisposition(name: string): string {
  const fallback = name.replace(/[^\x20-\x7e]|["\\]/g, '_') || 'attachment';
  return `attachment; filename="${fallback}"; filename*=UTF-8''${encodeURIComponent(name)}`;
}

function sendError(res: any, status: number, code: string, message: string): void {
  res.status(status).set('Cache-Control', 'no-store').json({ error: { code, message } });
}

export function mountAttachmentRoutes(app: any, config: AppConfig): void {
  if (!attachmentLinksAvailable(config)) return;

  for (const routePath of mountedRoutePaths(config, '/attachments/:ticket')) {
    app.get(routePath, async (req: any, res: any) => {
      const link = attachmentLinkStore.redeem(String(req.params.ticket ?? ''));
      if (!link) {
        attachmentLog('rejected');
        sendError(res, 410, 'link_expired', 'This download link is expired or was already used. Ask for a new one.');
        return;
      }

      let response: Response;
      try {
        response = await fetch(`${config.graphBaseUrl}${link.graphPath}`, {
          headers: { Authorization: `Bearer ${link.accessToken}`, Accept: '*/*' },
          signal: AbortSignal.timeout(Math.max(config.graphResilience.timeoutMs, 120_000))
        });
      } catch (error) {
        attachmentLog('graph_unreachable', { error: error instanceof Error ? error.name : 'unknown' });
        sendError(res, 502, 'graph_unreachable', 'Microsoft Graph could not be reached.');
        return;
      }

      if (!response.ok || !response.body) {
        await response.body?.cancel().catch(() => undefined);
        attachmentLog('graph_error', { status: response.status });
        if (response.status === 401) {
          sendError(res, 410, 'link_expired', 'The sign-in behind this link expired. Ask for a new one.');
        } else if (response.status === 403 || response.status === 404) {
          sendError(res, response.status, 'not_available', 'The attachment is not available to the signed-in user.');
        } else {
          sendError(res, 502, 'graph_error', `Microsoft Graph answered HTTP ${response.status}.`);
        }
        return;
      }

      res.status(200).set({
        'Content-Type': link.contentType || response.headers.get('content-type') || 'application/octet-stream',
        'Content-Disposition': contentDisposition(link.name),
        'Cache-Control': 'no-store',
        'X-Content-Type-Options': 'nosniff',
        'Content-Security-Policy': "default-src 'none'; sandbox",
        'Referrer-Policy': 'no-referrer'
      });
      const length = response.headers.get('content-length');
      if (length) res.set('Content-Length', length);

      const body = Readable.fromWeb(response.body as unknown as WebReadableStream<Uint8Array>);
      body.on('error', () => res.destroy());
      body.pipe(res);
      res.on('finish', () => attachmentLog('redeemed', { bytes: length ? Number(length) : undefined }));
    });
  }
}
