/**
 * Office / web paste cleanup
 *
 * Content pasted from word processors and web pages arrives wrapped in a large
 * amount of producer-specific cruft: conditional comments, XML processing
 * instructions, namespaced markup (`<o:p>`, `<w:sdt>`, smart tags), `mso-*`
 * style declarations, and empty spans. None of it maps to the editor schema,
 * and some of it confuses the browser HTML parser that ProseMirror's clipboard
 * pipeline relies on. Producer class names (`MsoNormal`, ...) are deliberately
 * kept so the downstream style inliner can match `<style>` rules against them.
 *
 * `cleanPastedHtml` normalizes the raw clipboard HTML string into something the
 * schema's `parseDOM` rules can read cleanly. It is a pure string transform so
 * it runs in the editor and in tests without a DOM, and it is deliberately
 * conservative: it strips producer metadata but never rewrites visible text
 * (curly quotes, non-Latin scripts, and whitespace between words are left
 * untouched, since this editor targets an international, typography-sensitive
 * audience).
 */

import { stripXmlDeclarations } from "../../../utils/stripXmlDeclarations";

/**
 * Remove every HTML comment, including downlevel conditional comments
 * (`<!--[if gte mso 9]> ... <![endif]-->`) that carry the Office `<xml>` island,
 * while leaving the contents of `<style>` elements untouched.
 *
 * The `<style>` carve-out matters because legacy pasted HTML hides stylesheet
 * CSS inside comment delimiters (`<style><!-- .c { ... } --></style>`, the old
 * CDATA-hiding trick). This cleanup runs before the style inliner and must hand
 * it an intact stylesheet, so a blanket comment strip would drop class rules
 * that used to be inlined. Style blocks are copied through verbatim.
 *
 * A single linear scan rather than a regex: clipboard HTML is untrusted, and a
 * lazy `<!--[\s\S]*?-->` against the multi-character `-->` terminator backtracks
 * polynomially on hostile input. The scan also drops an unterminated comment
 * through end-of-string so no stray `<!--` can survive.
 */
function stripHtmlComments(html: string): string {
  const lower = html.toLowerCase();
  let result = "";
  let cursor = 0;

  while (cursor < html.length) {
    const commentStart = html.indexOf("<!--", cursor);
    if (commentStart === -1) {
      result += html.slice(cursor);
      break;
    }

    // A `<style>` element that opens before the next comment is copied through
    // verbatim (delimiters and all) so the downstream inliner still sees its CSS.
    const styleStart = lower.indexOf("<style", cursor);
    if (styleStart !== -1 && styleStart < commentStart) {
      const openTagEnd = html.indexOf(">", styleStart);
      const closeStart = openTagEnd === -1 ? -1 : lower.indexOf("</style", openTagEnd + 1);
      const closeTagEnd = closeStart === -1 ? -1 : html.indexOf(">", closeStart);
      if (closeTagEnd === -1) {
        // Malformed/unterminated <style>: keep the remainder as-is and stop.
        result += html.slice(cursor);
        break;
      }
      result += html.slice(cursor, closeTagEnd + 1);
      cursor = closeTagEnd + 1;
      continue;
    }

    result += html.slice(cursor, commentStart);
    const commentEnd = html.indexOf("-->", commentStart + 4);
    if (commentEnd === -1) {
      break;
    }

    cursor = commentEnd + 3;
  }

  return result;
}

/**
 * Drop `mso-*` declarations from every inline `style` attribute while keeping
 * legitimate CSS. Operates per attribute (bounded `[^"]*` / `[^']*` inside the
 * quotes) so it can never run past the attribute into unrelated markup. An
 * attribute left with no non-`mso-*` declarations is removed entirely.
 */
function stripMsoStyles(html: string): string {
  // Capture the leading whitespace too so dropping an emptied attribute does not
  // leave a dangling space (`<span >`).
  return html.replace(
    /(\s+)style=(?:"([^"]*)"|'([^']*)')/gi,
    (_match, ws: string, dq: string | undefined, sq: string | undefined) => {
      const raw = dq ?? sq ?? "";
      const quote = dq === undefined ? "'" : '"';
      const kept = splitStyleDeclarations(raw)
        .map((declaration) => declaration.trim())
        .filter((declaration) => declaration.length > 0 && !/^mso-/i.test(declaration))
        .join("; ");
      return kept ? `${ws}style=${quote}${kept}${quote}` : "";
    },
  );
}

/**
 * Split a CSS `style` value into declarations on top-level `;` only.
 *
 * A naive `split(";")` corrupts values that legitimately contain a semicolon
 * inside a function or a quoted string (e.g. `background:url("data:...;base64,...")`).
 * This scan ignores `;` while inside parentheses or a quoted string, so such
 * declarations stay intact.
 */
function splitStyleDeclarations(value: string): string[] {
  const declarations: string[] = [];
  let current = "";
  let parenDepth = 0;
  let quote: '"' | "'" | null = null;

  for (const char of value) {
    if (quote !== null) {
      current += char;
      if (char === quote) {
        quote = null;
      }
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
    } else if (char === "(") {
      parenDepth++;
    } else if (char === ")") {
      parenDepth = Math.max(0, parenDepth - 1);
    } else if (char === ";" && parenDepth === 0) {
      declarations.push(current);
      current = "";
      continue;
    }
    current += char;
  }

  if (current) {
    declarations.push(current);
  }
  return declarations;
}

// Producer-only class names (`MsoNormal`, `MsoListParagraph`, ...) are left in
// place on purpose. The downstream style inliner matches them against `<style>`
// rules to inline real Word formatting (e.g. `.MsoNormal { font-size: 18pt }`),
// and the schema's `parseDOM` keys off tags, marks, and specific attributes —
// never arbitrary class names — so leftover producer classes are dropped on
// parse and cause no harm. Stripping them here (before the inliner runs) would
// break class-based paste, so it is intentionally not done.

// Matches the local name and attributes after a tag's prefix, treating quoted
// attribute values as opaque so a `>` inside a value (e.g. `title="a>b"`) does
// not truncate the tag mid-attribute. Stays linear: each alternative begins on
// a distinct character (`"`, `'`, or any other non-`>` char).
const TAG_TAIL = `(?:"[^"]*"|'[^']*'|[^>"'])*`;

// Word/Office/VML/OMML namespaced tags and smart tags (`<o:p>`, `<w:sdt>`,
// `<v:shape>`, `<m:oMath>`, `<st1:place>`). Tag-only removal keeps any inner
// text so wrapped user content is never deleted.
const NAMESPACED_TAG = new RegExp(`<\\/?(?:[owvm]|st\\d+):${TAG_TAIL}>`, "gi");
const STRAY_XML_TAG = new RegExp(`<\\/?xml${TAG_TAIL}>`, "gi");
const NOISE_TAG = new RegExp(`<\\/?(?:font|meta|link)${TAG_TAIL}>`, "gi");
const EMPTY_SPAN = new RegExp(`<span(?:\\s${TAG_TAIL})?><\\/span>`, "gi");
const MAX_EMPTY_SPAN_PASSES = 5;

// ProseMirror's own clipboard serializer wraps a copied slice's top element
// with `data-pm-slice="<openStart> <openEnd> <context>"` (see
// prosemirror-view's `serializeForClipboard`). Its presence is the only
// signal available at this string-transform layer that the incoming HTML
// came from a ProseMirror editor's own copy rather than an arbitrary
// external DOM (another app, a browser page, a hand-crafted paste).
const PM_SLICE_MARKER = /\bdata-pm-slice\s*=/i;

// `data-docx-textbox-anchor` is an internal reconstruction marker
// (`TextBoxAnchorExtension`) that `fromProseDoc` uses to relocate a
// paragraph's sibling text box on save, registering the first anchor it
// sees for a given id. It should never originate from outside a
// ProseMirror-serialized folio slice: an external page could plant a span
// carrying an id that collides with — and hijacks — a real text box's
// anchor (relocating it, or wrapping it in an attacker-controlled
// hyperlink). Strip the attribute so a matching span parses as an inert,
// zero-width `<span>` instead of a `textBoxAnchor` node.
const TEXTBOX_ANCHOR_ATTR = /\s+data-docx-textbox-anchor(?:="[^"]*"|='[^']*')?/gi;

/**
 * Remove the `data-docx-textbox-anchor` marker from HTML that did not come
 * from a ProseMirror clipboard slice (see {@link PM_SLICE_MARKER}). Internal
 * copy/paste of a real text box carries the marker HTML unmodified so it
 * keeps working; anything else has the marker stripped defensively.
 */
function stripForeignTextBoxAnchors(html: string, originalHtml: string): string {
  if (PM_SLICE_MARKER.test(originalHtml)) {
    return html;
  }
  return html.replace(TEXTBOX_ANCHOR_ATTR, "");
}

/**
 * Remove empty spans left behind after stripping `mso-*` styles. Only truly
 * empty spans are removed (whitespace-only spans are kept so word-separating
 * spacer runs never collapse two words together). Repeated a bounded number of
 * times to unwrap nested empties (`<span><span></span></span>`).
 */
function stripEmptySpans(html: string): string {
  let current = html;
  for (let pass = 0; pass < MAX_EMPTY_SPAN_PASSES; pass++) {
    const next = current.replace(EMPTY_SPAN, "");
    if (next === current) {
      break;
    }
    current = next;
  }
  return current;
}

/**
 * Strip Office/web producer cruft from a raw clipboard HTML string.
 *
 * Best-effort and non-throwing: any unexpected failure returns the original
 * markup so a paste degrades to the browser default rather than losing content.
 * `<style>` blocks are intentionally left in place for the downstream style
 * inliner, which resolves class-based CSS before the schema parser runs.
 */
export function cleanPastedHtml(html: string): string {
  if (!html) {
    return html;
  }

  try {
    let cleaned = stripHtmlComments(html);
    cleaned = stripXmlDeclarations(cleaned);
    cleaned = cleaned.replace(NAMESPACED_TAG, "");
    cleaned = cleaned.replace(STRAY_XML_TAG, "");
    cleaned = cleaned.replace(NOISE_TAG, "");
    cleaned = stripMsoStyles(cleaned);
    cleaned = stripEmptySpans(cleaned);
    cleaned = stripForeignTextBoxAnchors(cleaned, html);
    return cleaned.trim();
  } catch {
    return html;
  }
}
