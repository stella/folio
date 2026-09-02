const ALLOWED_URL_PROTOCOLS = new Set(["http:", "https:", "mailto:", "tel:"]);

/**
 * Keep only http(s), mailto, and tel URLs, normalised through the URL parser.
 * Anything else (javascript:, data:, relative paths, malformed input) drops to
 * `undefined` so a markdown link degrades to its text instead of carrying an
 * executable target into the document.
 */
export const sanitizeExternalUrl = (rawUrl: string | undefined): string | undefined => {
  if (!rawUrl) {
    return undefined;
  }

  const trimmed = rawUrl.trim();
  if (!trimmed) {
    return undefined;
  }

  const parsed = parseUrl(trimmed);
  if (!parsed || !ALLOWED_URL_PROTOCOLS.has(parsed.protocol)) {
    return undefined;
  }
  if (
    (parsed.protocol === "mailto:" || parsed.protocol === "tel:") &&
    parsed.pathname.trim() === ""
  ) {
    return undefined;
  }
  return parsed.href;
};

// `URL` throws on malformed input; an unparseable link is a normal outcome
// here (the link renders as plain text), not a failure to propagate.
const parseUrl = (value: string): URL | null => {
  try {
    return new URL(value);
  } catch {
    return null;
  }
};

const hasUnsafeAnchorCharacter = (anchor: string): boolean => {
  for (const char of anchor) {
    const codePoint = char.codePointAt(0) ?? 0;
    if (codePoint <= 0x20 || codePoint === 0x7f || char.trim() === "") {
      return true;
    }
  }
  return false;
};

/** A `#anchor` stays a document-internal target; anything else must pass {@link sanitizeExternalUrl}. */
export const sanitizeMarkdownHref = (rawHref: string): string | undefined => {
  const trimmed = rawHref.trim();
  if (!trimmed) {
    return undefined;
  }

  if (trimmed.startsWith("#")) {
    const anchor = trimmed.slice(1);
    if (!anchor || hasUnsafeAnchorCharacter(anchor)) {
      return undefined;
    }
    return `#${anchor}`;
  }

  return sanitizeExternalUrl(trimmed);
};
