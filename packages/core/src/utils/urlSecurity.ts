import { sanitizeExternalUrl } from "@stll/docx-core";

const ALLOWED_TARGETS = new Set(["_blank", "_self", "_parent", "_top"]);

// One allowlist for every link that enters a document, whether typed in the
// editor or parsed from markdown: the rule lives in docx-core.
export { sanitizeExternalUrl };

export function normalizeUserUrl(rawUrl: string): string {
  const trimmed = rawUrl.trim();
  if (!trimmed) {
    return "";
  }

  const normalized = trimmed.toLowerCase();
  if (normalized.startsWith("mailto:") || normalized.startsWith("tel:")) {
    return sanitizeExternalUrl(trimmed) ?? "";
  }

  if (normalized.startsWith("http://") || normalized.startsWith("https://")) {
    return sanitizeExternalUrl(trimmed) ?? "";
  }

  // Protocol-less input with a colon is either host:port, bracketed IPv6,
  // or an unsupported scheme. Reject the scheme-like forms before adding https.
  const colonIndex = trimmed.indexOf(":");
  const firstPathSeparatorIndex = findFirstPathSeparatorIndex(trimmed);
  const colonIsInAuthority =
    colonIndex > 0 &&
    !trimmed.startsWith("[") &&
    (firstPathSeparatorIndex === -1 || colonIndex < firstPathSeparatorIndex);
  if (colonIsInAuthority) {
    const suffix = trimmed.slice(colonIndex + 1);
    if (suffix.startsWith("//") || !startsWithPortSuffix(suffix)) {
      return "";
    }
  }

  const withProtocol = `https://${trimmed}`;
  return sanitizeExternalUrl(withProtocol) ?? "";
}

export function isAllowedUserUrl(rawUrl: string): boolean {
  return normalizeUserUrl(rawUrl) !== "";
}

/**
 * The `target` and `rel` one authored frame may produce on a DOM anchor.
 *
 * A `w:tgtFrame` is a frame name, and a named frame addresses another browsing
 * context: honouring an arbitrary one would let a document reach a frame it did
 * not create. Every anchor and every navigation folio produces goes through
 * here, so the clamp and the `rel` that must accompany `_blank` are decided
 * once. The authored name itself stays in the model and is written back on
 * save; clamping it at parse rewrote the document.
 */
const ANCHOR_REL = "noopener noreferrer";

export type AnchorTargetAttrs = { target: string; rel: string };

export function anchorTargetAttrs(authoredFrame: string | undefined): AnchorTargetAttrs {
  return {
    target: authoredFrame && ALLOWED_TARGETS.has(authoredFrame) ? authoredFrame : "_blank",
    rel: ANCHOR_REL,
  };
}

/** `window.open` features carrying the same guarantees as {@link ANCHOR_REL}. */
export function windowFeaturesFor({ rel }: AnchorTargetAttrs): string {
  return rel.replaceAll(" ", ",");
}

function findFirstPathSeparatorIndex(value: string): number {
  const indexes = [value.indexOf("/"), value.indexOf("?"), value.indexOf("#")]
    .filter((index) => index >= 0)
    .toSorted((a, b) => a - b);
  return indexes.at(0) ?? -1;
}

function startsWithPortSuffix(value: string): boolean {
  if (value.length === 0) {
    return false;
  }

  let index = 0;
  while (index < value.length) {
    const char = value.codePointAt(index) ?? 0;
    if (char < 48 || char > 57) {
      break;
    }
    index += 1;
  }

  return index > 0 && (index === value.length || "/?#".includes(value[index]!));
}
