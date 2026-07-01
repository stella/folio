/**
 * Office / web paste cleanup
 *
 * Content pasted from word processors and web pages arrives wrapped in a large
 * amount of producer-specific cruft: conditional comments, XML processing
 * instructions, namespaced markup (`<o:p>`, `<w:sdt>`, smart tags), `mso-*`
 * style declarations, producer-only class names, and empty spans. None of it
 * maps to the editor schema, and some of it confuses the browser HTML parser
 * that ProseMirror's clipboard pipeline relies on.
 *
 * `cleanPastedHtml` normalizes the raw clipboard HTML string into something the
 * schema's `parseDOM` rules can read cleanly. It is a pure string transform so
 * it runs in the editor and in tests without a DOM, and it is deliberately
 * conservative: it strips producer metadata but never rewrites visible text
 * (curly quotes, non-Latin scripts, and whitespace between words are left
 * untouched, since this editor targets an international, typography-sensitive
 * audience).
 */

/**
 * Remove every HTML comment, including downlevel conditional comments
 * (`<!--[if gte mso 9]> ... <![endif]-->`) that carry the Office `<xml>` island.
 *
 * A single linear scan rather than a regex: clipboard HTML is untrusted, and a
 * lazy `<!--[\s\S]*?-->` against the multi-character `-->` terminator backtracks
 * polynomially on hostile input. The scan also drops an unterminated comment
 * through end-of-string so no stray `<!--` can survive.
 */
function stripHtmlComments(html: string): string {
  let result = "";
  let cursor = 0;

  while (cursor < html.length) {
    const start = html.indexOf("<!--", cursor);
    if (start === -1) {
      result += html.slice(cursor);
      break;
    }

    result += html.slice(cursor, start);
    const end = html.indexOf("-->", start + 4);
    if (end === -1) {
      break;
    }

    cursor = end + 3;
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

/**
 * Drop producer-only class tokens (`MsoNormal`, `MsoListParagraph`, ...) while
 * keeping any real class names. Bounded per attribute like {@link stripMsoStyles}.
 */
function stripMsoClasses(html: string): string {
  return html.replace(
    /(\s+)class=(?:"([^"]*)"|'([^']*)')/gi,
    (_match, ws: string, dq: string | undefined, sq: string | undefined) => {
      const raw = dq ?? sq ?? "";
      const quote = dq === undefined ? "'" : '"';
      const kept = raw
        .split(/\s+/)
        .filter((token) => token.length > 0 && !/^mso/i.test(token))
        .join(" ");
      return kept ? `${ws}class=${quote}${kept}${quote}` : "";
    },
  );
}

// Matches the local name and attributes after a tag's prefix, treating quoted
// attribute values as opaque so a `>` inside a value (e.g. `title="a>b"`) does
// not truncate the tag mid-attribute. Stays linear: each alternative begins on
// a distinct character (`"`, `'`, or any other non-`>` char).
const TAG_TAIL = `(?:"[^"]*"|'[^']*'|[^>"'])*`;

// Word/Office/VML/OMML namespaced tags and smart tags (`<o:p>`, `<w:sdt>`,
// `<v:shape>`, `<m:oMath>`, `<st1:place>`). Tag-only removal keeps any inner
// text so wrapped user content is never deleted.
const NAMESPACED_TAG = new RegExp(`<\\/?(?:[owvm]|st\\d+):${TAG_TAIL}>`, "gi");
const XML_PROCESSING_INSTRUCTION = new RegExp(`<\\?xml${TAG_TAIL}>`, "gi");
const STRAY_XML_TAG = new RegExp(`<\\/?xml${TAG_TAIL}>`, "gi");
const NOISE_TAG = new RegExp(`<\\/?(?:font|meta|link)${TAG_TAIL}>`, "gi");
const EMPTY_SPAN = new RegExp(`<span(?:\\s${TAG_TAIL})?><\\/span>`, "gi");
const MAX_EMPTY_SPAN_PASSES = 5;

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
    cleaned = cleaned.replace(XML_PROCESSING_INSTRUCTION, "");
    cleaned = cleaned.replace(NAMESPACED_TAG, "");
    cleaned = cleaned.replace(STRAY_XML_TAG, "");
    cleaned = cleaned.replace(NOISE_TAG, "");
    cleaned = stripMsoStyles(cleaned);
    cleaned = stripMsoClasses(cleaned);
    cleaned = stripEmptySpans(cleaned);
    return cleaned.trim();
  } catch {
    return html;
  }
}
