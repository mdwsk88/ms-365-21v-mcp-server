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

export function normalizeText(value: string): string {
  return value
    .replace(/\r\n?/g, '\n')
    .replace(/[\t\f\v ]+/g, ' ')
    .split('\n')
    .map(line => line.replace(/ {2,}/g, ' ').replace(/\s*\|\s*$/, '').trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
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
