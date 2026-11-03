const namedEntities: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
  ensp: ' ',
  emsp: ' ',
  thinsp: ' ',
  zwnj: '',
  zwj: '',
  shy: '',
  hellip: '…',
  mdash: '—',
  ndash: '–',
  lsquo: '‘',
  rsquo: '’',
  ldquo: '“',
  rdquo: '”',
  bull: '•',
  middot: '·',
  copy: '©',
  reg: '®',
  trade: '™'
};

export function decodeHtmlEntities(value: string): string {
  return value.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (match, entity: string) => {
    if (entity[0] === '#') {
      const codePoint =
        entity[1] === 'x' || entity[1] === 'X'
          ? Number.parseInt(entity.slice(2), 16)
          : Number.parseInt(entity.slice(1), 10);
      return Number.isFinite(codePoint) && codePoint > 0 && codePoint <= 0x10ffff
        ? String.fromCodePoint(codePoint)
        : match;
    }
    return namedEntities[entity.toLowerCase()] ?? match;
  });
}

/**
 * Converts Outlook and Teams HTML into readable plain text: keeps paragraph,
 * list and table structure, drops styles, scripts, comments and inline attachment markers.
 */
export function htmlToText(html: string): string {
  if (!html) return '';
  const text = html
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(script|style|head|title)\b[\s\S]*?<\/\1>/gi, '')
    .replace(/<attachment\b[^>]*>\s*<\/attachment>/gi, '')
    .replace(/<emoji\b[^>]*\balt="([^"]*)"[^>]*>(?:\s*<\/emoji>)?/gi, '$1')
    .replace(/<img\b[^>]*>/gi, '[image]')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<li\b[^>]*>/gi, '\n- ')
    .replace(/<\/li>/gi, '')
    .replace(/<\/(td|th)>/gi, ' | ')
    .replace(/<\/(p|div|tr|ul|ol|table|h[1-6]|blockquote|pre|codeblock|section|article|header|footer)>/gi, '\n')
    .replace(/<(p|div|table|h[1-6]|blockquote|pre|codeblock)\b[^>]*>/gi, '\n')
    .replace(/<[^>]+>/g, '');
  return normalizeText(decodeHtmlEntities(text));
}

const invisibleCharacters =
  /[­͏ᅟᅠ឴឵᠎​‎‏‪-‮⁠-⁤⁦-⁯ㅤ﻿ﾠ]/g;
// Zero-width joiners belong inside emoji and some scripts; next to spaces or each other they are padding.
const strayJoiners = /(?<=^|[\s‌‍])[‌‍]+|[‌‍]+(?=\s|$)/gm;
const spaceLikeCharacters = /[  -   ]/g;

/** Removes zero-width, bidi-control and padding characters, such as the preview padding of marketing mail. */
export function removeInvisibleCharacters(value: string): string {
  return value.replace(invisibleCharacters, '').replace(strayJoiners, '').replace(spaceLikeCharacters, ' ');
}

export function normalizeText(value: string): string {
  return removeInvisibleCharacters(value)
    .replace(/\r\n?/g, '\n')
    .replace(/[\t\f\v ]+/g, ' ')
    .split('\n')
    .map(line => line.replace(/ {2,}/g, ' ').replace(/\s*\|\s*$/, '').trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/** Text file content: unifies line endings and drops a byte-order mark, but keeps tabs and indentation. */
export function fileText(value: string): string {
  return value.replace(/^﻿/, '').replace(/\r\n?/g, '\n');
}

// Mail bodies

const linkCharacters = String.raw`[\w\-.~:/?#\[\]@!$&'()*+,;=%]`;
const linkPattern = new RegExp(String.raw`(<)?(https?://${linkCharacters}+)(>)?`, 'gi');
const trailingLink = new RegExp(String.raw`https?://${linkCharacters}+$`, 'i');
const safeLinksHost = /(^|\.)safelinks\.protection\.[a-z0-9.-]+$/i;
const trackingParameter = /^(utm_\w+|mc_[ce]id|fbclid|gclid|dclid|gbraid|wbraid|msclkid|yclid|igshid|_hsenc|_hsmi|mkt_tok|_ga|_gl)$/i;
// Links to Microsoft 365 content are kept whole however long: the agent may need to open them.
const workLinkHost =
  /(^|\.)(sharepoint\.(com|cn)|microsoftonline\.cn|teams\.microsoft\.com|teams\.live\.com|office\.com|office365\.cn|onedrive\.live\.com|1drv\.ms)$/i;
const longLinkChars = 500;
// How far back to look for the text before a link: enough for any link worth comparing, and it keeps
// the cleanup linear in the body length.
const lookBehindChars = 2100;

function decodeComponent(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

/**
 * Unwraps Microsoft Defender Safe Links (any cloud, including 21Vianet) and removes marketing
 * tracking parameters. Links that need neither come back exactly as written.
 */
export function cleanUrl(raw: string): string {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    return raw;
  }
  let changed = false;
  for (let depth = 0; depth < 3 && safeLinksHost.test(url.hostname); depth += 1) {
    const target = url.searchParams.get('url');
    if (!target || !/^https?:\/\//i.test(target)) break;
    try {
      url = new URL(target);
    } catch {
      break;
    }
    changed = true;
  }
  if (url.search) {
    const pairs = url.search.slice(1).split('&');
    const kept = pairs.filter(pair => !trackingParameter.test(decodeComponent(pair.split('=')[0] ?? '')));
    if (kept.length !== pairs.length) {
      url.search = kept.join('&');
      changed = true;
    }
  }
  return changed ? url.href : raw;
}

function comparableUrl(value: string): string {
  try {
    return new URL(value).href.replace(/\/$/, '');
  } catch {
    return value.replace(/\/$/, '');
  }
}

function count(value: string, character: string): number {
  return value.split(character).length - 1;
}

/** Splits sentence punctuation and unbalanced closing brackets off the end of a matched link. */
function splitTrailingPunctuation(link: string): [string, string] {
  let end = link.length;
  while (end > 0) {
    const last = link[end - 1];
    const head = link.slice(0, end);
    if (".,;:!?'*".includes(last)) end -= 1;
    else if (last === ')' && count(head, '(') < count(head, ')')) end -= 1;
    else if (last === ']' && count(head, '[') < count(head, ']')) end -= 1;
    else break;
  }
  return [link.slice(0, end), link.slice(end)];
}

/**
 * Makes a plain-text mail body cheaper to read without changing its meaning: unwraps Safe Links,
 * drops tracking parameters and repeated link targets, turns inline image references into [image],
 * and replaces links longer than 500 characters, other than Microsoft 365 links, with [link: host].
 */
export function cleanMailText(value: string): { text: string; linksOmitted: number } {
  let linksOmitted = 0;
  const shortened = (link: string): string | undefined => {
    if (link.length <= longLinkChars) return undefined;
    let host: string;
    try {
      host = new URL(link).hostname;
    } catch {
      return undefined;
    }
    if (workLinkHost.test(host)) return undefined;
    linksOmitted += 1;
    return `[link: ${host}]`;
  };

  const cleaned = value
    .replace(/\[cid:[^\]\s]*\]/gi, '[image]')
    .replace(/<mailto:([^>\s?]+)(?:\?[^>\s]*)?>/gi, (_match, address: string, offset: number, whole: string) => {
      const before = whole.slice(Math.max(0, offset - lookBehindChars), offset);
      if (before.trimEnd().toLowerCase().endsWith(decodeComponent(address).toLowerCase())) return '';
      return `${/\s$/.test(before) || !before ? '' : ' '}<${decodeComponent(address)}>`;
    })
    .replace(linkPattern, (_match, open: string | undefined, link: string, close: string | undefined, offset: number, whole: string) => {
      const before = whole.slice(Math.max(0, offset - lookBehindChars), offset);
      if (open && close) {
        // Outlook renders a hyperlink as "text<target>"; when the text is the same link, keep one copy.
        const target = cleanUrl(link);
        const shown = before.match(trailingLink)?.[0];
        if (shown && comparableUrl(cleanUrl(splitTrailingPunctuation(shown)[0])) === comparableUrl(target)) return '';
        const short = shortened(target);
        return short ? `${/\s$/.test(before) || !before ? '' : ' '}${short}` : `<${target}>`;
      }
      const [core, trailing] = splitTrailingPunctuation(link);
      const target = cleanUrl(core);
      return `${open ?? ''}${shortened(target) ?? target}${trailing}${close ?? ''}`;
    });
  return { text: normalizeText(cleaned), linksOmitted };
}

export function truncateText(value: string, maxChars: number): { text: string; truncated: boolean } {
  const characters = Array.from(value);
  if (characters.length <= maxChars) return { text: value, truncated: false };
  return { text: `${characters.slice(0, Math.max(0, maxChars - 1)).join('')}…`, truncated: true };
}

export function preview(value: string | undefined, maxChars = 280): string | undefined {
  if (!value) return undefined;
  const single = normalizeText(value).replace(/\s*\n\s*/g, ' ');
  return single ? truncateText(single, maxChars).text : undefined;
}

const droppedGraphKeys = new Set(['@odata.context', '@odata.etag']);

/**
 * Removes Graph transport metadata and null values so tool results spend
 * model context on business data. Arrays keep their positions.
 */
export function pruneGraphPayload(value: unknown, depth = 0): unknown {
  if (depth > 32) return value;
  if (Array.isArray(value)) return value.map(item => pruneGraphPayload(item, depth + 1));
  if (!value || typeof value !== 'object') return value;
  const output: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    if (child === null || child === undefined || droppedGraphKeys.has(key)) continue;
    output[key] = pruneGraphPayload(child, depth + 1);
  }
  return output;
}
