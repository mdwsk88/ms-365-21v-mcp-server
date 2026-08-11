import type { AppConfig } from './config.js';

function normalizeRoutePath(value: string): string {
  const withLeadingSlash = value.startsWith('/') ? value : `/${value}`;
  const stripped = withLeadingSlash.replace(/\/+$/, '');
  return stripped || '/';
}

export function mountedRoutePaths(config: AppConfig, pathname: string): string[] {
  const routePath = normalizeRoutePath(pathname);
  const paths = new Set([routePath]);
  const prefix = config.publicPathPrefix;

  if (prefix && routePath !== prefix && !routePath.startsWith(`${prefix}/`)) {
    paths.add(routePath === '/' ? prefix : `${prefix}${routePath}`);
  }

  return [...paths];
}
