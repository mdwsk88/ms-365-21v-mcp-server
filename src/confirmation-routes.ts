import type { AppConfig } from './config.js';
import { hasConfirmationPolicy } from './confirmation.js';
import {
  browserConfirmationStore,
  type BrowserConfirmationView,
  type BrowserOperationPreview,
  type PreviewField
} from './browser-confirmation.js';
import { mountedRoutePaths } from './http-routes.js';

function escapeHtml(value: unknown): string {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function cookieName(confirmToken: string): string {
  return `mcp_confirmation_${confirmToken.slice(-12)}`;
}

function cookieValue(header: string | undefined, name: string): string | undefined {
  if (!header) return undefined;
  for (const part of header.split(';')) {
    const separator = part.indexOf('=');
    if (separator < 0 || part.slice(0, separator).trim() !== name) continue;
    return decodeURIComponent(part.slice(separator + 1).trim());
  }
  return undefined;
}

function approvalCookie(view: BrowserConfirmationView, confirmToken: string, browserNonce: string): string {
  const secure = new URL(view.approvalUrl).protocol === 'https:' ? '; Secure' : '';
  const path = new URL(view.approvalUrl).pathname;
  return `${cookieName(confirmToken)}=${encodeURIComponent(browserNonce)}; Path=${path}; HttpOnly; SameSite=Strict; Max-Age=${view.expiresIn}${secure}`;
}

function expiredCookie(view: BrowserConfirmationView | undefined, confirmToken: string): string {
  const path = view ? new URL(view.approvalUrl).pathname : `/confirm/${encodeURIComponent(confirmToken)}`;
  const secure = view && new URL(view.approvalUrl).protocol === 'https:' ? '; Secure' : '';
  return `${cookieName(confirmToken)}=; Path=${path}; HttpOnly; SameSite=Strict; Max-Age=0${secure}`;
}

function secureHtml(res: any): void {
  res.set('Cache-Control', 'no-store');
  res.set(
    'Content-Security-Policy',
    "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'"
  );
  res.set('Referrer-Policy', 'no-referrer');
  res.set('X-Content-Type-Options', 'nosniff');
  res.set('X-Frame-Options', 'DENY');
  res.set('Permissions-Policy', 'camera=(), microphone=(), geolocation=()');
}

function page(title: string, body: string, destructive = false): string {
  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${escapeHtml(title)}</title>
  <style>
    :root{color-scheme:light;--ink:#17202a;--muted:#5e6b78;--line:#dce2e8;--surface:#fff;--page:#f3f5f7;--accent:#1261a0;--accent-hover:#0d4d80;--danger:#b42318;--danger-soft:#fff1f0;--ok:#18794e;--ok-soft:#edf8f3}
    *{box-sizing:border-box}body{margin:0;background:var(--page);color:var(--ink);font:15px/1.55 system-ui,-apple-system,"Segoe UI","Microsoft YaHei",sans-serif}
    main{width:min(820px,calc(100% - 32px));margin:36px auto 64px;background:var(--surface);border:1px solid var(--line);border-radius:8px;box-shadow:0 8px 28px rgba(23,32,42,.07);overflow:hidden}
    header{padding:26px 30px 22px;border-bottom:1px solid var(--line)}.brand{color:var(--muted);font-size:13px;font-weight:650;margin-bottom:8px}.title-row{display:flex;gap:13px;align-items:flex-start}.mark{display:grid;place-items:center;flex:0 0 34px;height:34px;border-radius:50%;background:${destructive ? 'var(--danger-soft)' : '#eaf3fb'};color:${destructive ? 'var(--danger)' : 'var(--accent)'};font-size:20px;font-weight:750}h1{font-size:24px;line-height:1.3;margin:1px 0 5px}.lead{color:var(--muted);margin:0}
    .content-wrap{padding:24px 30px 30px}.summary{display:flex;justify-content:space-between;gap:16px;align-items:center;margin-bottom:18px}.operation{font-size:13px;color:var(--muted)}code{font-family:ui-monospace,SFMono-Regular,Consolas,monospace;background:#f0f3f6;border-radius:4px;padding:2px 6px;color:#334155;overflow-wrap:anywhere}
    dl{margin:0;border-top:1px solid var(--line)}.field{display:grid;grid-template-columns:180px minmax(0,1fr);gap:20px;padding:13px 0;border-bottom:1px solid var(--line)}dt{color:var(--muted);font-weight:600}dd{margin:0;white-space:pre-wrap;overflow-wrap:anywhere}
    .body-section{margin-top:22px}.section-label{font-weight:700;margin-bottom:9px}.body-preview{margin:0;min-height:120px;max-height:420px;overflow:auto;white-space:pre-wrap;overflow-wrap:anywhere;background:#f6f8fa;border:1px solid var(--line);border-radius:6px;padding:16px;font:14px/1.65 ui-monospace,SFMono-Regular,Consolas,monospace}
    .notice{margin-top:22px;padding:13px 15px;background:#f5f8fb;border-left:4px solid var(--accent);color:#34495e}.notice.danger{background:var(--danger-soft);border-color:var(--danger);color:#7a271a}.expiry{color:var(--muted);font-size:13px;margin:12px 0 0}
    .actions{display:flex;gap:12px;margin-top:24px}.actions form{margin:0}button{border:1px solid transparent;border-radius:6px;padding:11px 18px;font:inherit;font-weight:700;cursor:pointer}.approve{background:${destructive ? 'var(--danger)' : 'var(--accent)'};color:#fff}.approve:hover{background:${destructive ? '#912018' : 'var(--accent-hover)' }}.reject{background:#fff;border-color:#b8c2cc;color:#253342}.reject:hover{background:#f3f5f7}
    .status{padding:30px}.status-box{margin-top:18px;padding:14px 16px;border-left:4px solid var(--ok);background:var(--ok-soft)}.status-box.error{border-color:var(--danger);background:var(--danger-soft)}
    @media(max-width:620px){main{width:100%;margin:0;min-height:100vh;border:0;border-radius:0;box-shadow:none}header,.content-wrap,.status{padding-left:20px;padding-right:20px}.field{grid-template-columns:1fr;gap:4px}.summary{align-items:flex-start;flex-direction:column}.actions{flex-direction:column}.actions form,.actions button{width:100%}}
  </style>
</head>
<body><main>${body}</main></body>
</html>`;
}

function renderField(field: PreviewField): string {
  return `<div class="field"><dt>${escapeHtml(field.label)}</dt><dd>${escapeHtml(field.value)}</dd></div>`;
}

function reviewPage(view: BrowserConfirmationView): string {
  const preview: BrowserOperationPreview = view.preview;
  const path = new URL(view.approvalUrl).pathname;
  const content = preview.content
    ? `<section class="body-section"><div class="section-label">${escapeHtml(preview.content.label)}</div><pre class="body-preview">${escapeHtml(preview.content.value)}</pre></section>`
    : '';
  const warning = preview.destructive
    ? '这是删除操作。批准后，客户端继续执行时可能无法恢复。'
    : '批准后，只有同一登录用户回到 MCP 客户端继续操作时才会执行。';
  return page(
    `确认：${preview.title}`,
    `<header>
      <div class="brand">21V Microsoft 365 MCP</div>
      <div class="title-row"><div class="mark">!</div><div><h1>${escapeHtml(preview.title)}</h1><p class="lead">请核对以下目标和内容，然后决定是否允许执行。</p></div></div>
    </header>
    <div class="content-wrap">
      <div class="summary"><strong>操作预览</strong><div class="operation">工具 <code>${escapeHtml(preview.toolName)}</code></div></div>
      <dl>${preview.fields.map(renderField).join('')}</dl>
      ${content}
      <div class="notice${preview.destructive ? ' danger' : ''}">${escapeHtml(warning)}</div>
      <p class="expiry">本次确认约 ${view.expiresIn} 秒后失效，批准后只能执行一次。</p>
      <div class="actions">
        <form method="post" action="${escapeHtml(path)}/approve"><button class="approve" type="submit">${escapeHtml(preview.approveLabel)}</button></form>
        <form method="post" action="${escapeHtml(path)}/reject"><button class="reject" type="submit">拒绝并取消</button></form>
      </div>
    </div>`,
    preview.destructive
  );
}

function statusPage(title: string, message: string, error = false): string {
  return page(
    title,
    `<div class="status"><div class="brand">21V Microsoft 365 MCP</div><h1>${escapeHtml(title)}</h1><div class="status-box${error ? ' error' : ''}">${escapeHtml(message)}</div></div>`,
    error
  );
}

export function mountConfirmationRoutes(app: any, config: AppConfig): void {
  if (!hasConfirmationPolicy(config)) return;

  for (const routePath of mountedRoutePaths(config, '/confirm/:confirmToken')) {
    app.get(routePath, (req: any, res: any) => {
      secureHtml(res);
      const confirmToken = String(req.params.confirmToken ?? '');
      const preparation = browserConfirmationStore.prepareBrowserApproval(confirmToken);
      if (!preparation) {
        res.status(410).type('html').send(statusPage('确认已失效', '该确认不存在、已过期、已拒绝或已经执行。', true));
        return;
      }
      if (preparation.view.approved) {
        res.type('html').send(statusPage('操作已批准', '请返回 MCP 客户端并告诉 Agent 继续执行。'));
        return;
      }
      res.set('Set-Cookie', approvalCookie(preparation.view, confirmToken, preparation.browserNonce!));
      res.type('html').send(reviewPage(preparation.view));
    });
  }

  for (const routePath of mountedRoutePaths(config, '/confirm/:confirmToken/approve')) {
    app.post(routePath, (req: any, res: any) => {
      secureHtml(res);
      const confirmToken = String(req.params.confirmToken ?? '');
      const browserNonce = cookieValue(req.headers.cookie, cookieName(confirmToken));
      const view = browserNonce
        ? browserConfirmationStore.approveFromBrowser(confirmToken, browserNonce)
        : undefined;
      if (!view) {
        res.status(403).type('html').send(statusPage('无法批准', '确认已失效或浏览器校验失败，请返回客户端重新发起操作。', true));
        return;
      }
      res.set('Set-Cookie', expiredCookie(view, confirmToken));
      res.type('html').send(statusPage('操作已批准', '预览已经批准。请返回 MCP 客户端并告诉 Agent 继续执行。'));
    });
  }

  for (const routePath of mountedRoutePaths(config, '/confirm/:confirmToken/reject')) {
    app.post(routePath, (req: any, res: any) => {
      secureHtml(res);
      const confirmToken = String(req.params.confirmToken ?? '');
      const browserNonce = cookieValue(req.headers.cookie, cookieName(confirmToken));
      const view = browserNonce
        ? browserConfirmationStore.rejectFromBrowser(confirmToken, browserNonce)
        : undefined;
      if (!view) {
        res.status(403).type('html').send(statusPage('无法拒绝', '确认已失效或浏览器校验失败。', true));
        return;
      }
      res.set('Set-Cookie', expiredCookie(view, confirmToken));
      res.type('html').send(statusPage('操作已取消', '该操作不会执行，可以关闭此页面。'));
    });
  }
}
