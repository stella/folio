/**
 * Markdown escaping. Only characters that would change the parsed structure
 * get escaped. Context-sensitive: pipes only matter inside table cells.
 *
 * Ported from eigenpal/docx-editor PR #595 (`@eigenpal/docx-editor-core/markdown`).
 */

/**
 * Escape only characters that would change markdown structure mid-text:
 *   - backslash, backtick, asterisk, brackets, tilde (always)
 *   - underscore: unless between two letters or digits (emphasis trigger)
 *   - angle brackets: only when they look like a tag/autolink
 *
 * Characters like `.` `-` `+` `#` `(` `)` `!` are only meaningful at line
 * starts or in specific adjacency; we let the block-level renderer decide
 * when to escape those.
 */
export function escapeInline(text: string): string {
  // A tilde pair reads as strikethrough, a single one too (GFM).
  let out = text.replace(/(?<ch>[\\`*[\]~])/gu, "\\$<ch>");
  // Only an underscore between two letters or digits cannot open or close
  // emphasis (`snake_case`); any other one is escaped, once.
  out = out.replace(/(?<![\p{L}\p{N}])_|_(?![\p{L}\p{N}])/gu, "\\_");
  // Angle brackets only escaped when they form a plausible HTML tag — an
  // alphabetic name immediately followed by attributes/end. Prose like `x < y`
  // and `i < n` stays untouched.
  out = out.replace(/<(?<tag>\/?[A-Za-z][\w-]*)(?=[\s/>])/gu, "\\<$<tag>");
  return out;
}

/**
 * Cells use `|` as the column separator. The input is already markdown (each
 * cell is a rendered paragraph that ran through {@link escapeInline}), so a
 * backslash escape would interact with existing escape sequences — encode the
 * pipe as the HTML entity instead, which GFM renders as a literal `|` without
 * treating it as a delimiter. Newlines become `<br>` to keep the row on one
 * line.
 */
export function escapeTableCell(text: string): string {
  return inlineBreaksToHtml(text.replace(/\|/gu, "&#124;"));
}

const TWO_SPACE_BREAK = "  \n";
const BACKSLASH_BREAK = "\\\n";

/** A blank as a character reference, which no line-edge rule strips. */
const BLANK_REFERENCES: Record<string, string> = { " ": "&#32;", "\t": "&#9;" };
const blankReferences = (blanks: string): string =>
  blanks.replace(/[ \t]/gu, (blank) => BLANK_REFERENCES[blank] ?? blank);

/**
 * CommonMark strips the blanks at the start of a line and takes the blanks
 * before a line ending as the hard break's own syntax, so text blanks next to
 * an inline break would be lost on reading. They are written as character
 * references instead, which read back as the blanks they are.
 */
export const protectBreakBlanks = (text: string): string => {
  const lines = text.split(TWO_SPACE_BREAK);
  let out = "";
  lines.forEach((line, index) => {
    const last = index === lines.length - 1;
    const start = index === 0 ? 0 : (/^[ \t]*/u.exec(line)?.[0].length ?? 0);
    const end = last
      ? line.length
      : line.length - (/[ \t]*$/u.exec(line)?.[0].length ?? 0);
    // A line of blanks alone sits beside a break on either side.
    out +=
      end <= start
        ? blankReferences(line)
        : blankReferences(line.slice(0, start)) +
          line.slice(start, end) +
          blankReferences(line.slice(end));
    if (last) return;
    // An empty line before a break would read as a blank line, which ends the
    // paragraph: its break takes the backslash form, which keeps it inside.
    out += line.length === 0 ? BACKSLASH_BREAK : TWO_SPACE_BREAK;
  });
  return out;
};

/** Single-line Markdown blocks express inline breaks without their two-space syntax. */
export const inlineBreaksToHtml = (text: string): string =>
  text
    .split(/\r?\n/u)
    .map((line, index, lines) => {
      if (index === lines.length - 1) return line;
      // A lone backslash is a break's own syntax: text backslashes are escaped.
      if (line === "\\") return "";
      let end = line.length;
      while (end > 0 && line.at(end - 1) === " ") end--;
      return line.slice(0, end);
    })
    .join("<br>");

/**
 * Hyperlink URLs in inline form need parens balanced; we URL-encode the
 * problematic characters rather than escape, so the link still resolves.
 * `encodeURIComponent` leaves `(` and `)` untouched, so an unbalanced paren in
 * a DOCX URL (e.g. `https://example.com/a)b`) would close the inline
 * `[text](url)` destination early; encode the parens explicitly.
 */
const LINK_URL_ENCODED: Record<string, string> = { "(": "%28", ")": "%29" };

export function escapeLinkUrl(url: string): string {
  return url.replace(
    /[()<>"\s]/gu,
    (ch) => LINK_URL_ENCODED[ch] ?? encodeURIComponent(ch),
  );
}

/**
 * Alt text inside `![alt](url)`. Raw (un-escaped) text, so escape the backslash
 * first, then the brackets that would confuse the link parser, and collapse
 * newlines so the alt stays single-line.
 */
export function escapeAltText(text: string): string {
  return text.replace(/(?<ch>[\\[\]])/gu, "\\$<ch>").replace(/\r?\n/gu, " ");
}

/**
 * A plain paragraph whose visible text begins with markdown block syntax (e.g.
 * `# Not a heading`, `- value`, `1. value`, `> quote`) would be reclassified as
 * a heading/list/blockquote on re-parse, even though Word carries no matching
 * style/list metadata. Escape the leading marker so literal text round-trips.
 */
export function escapeLeadingBlockMarker(text: string): string {
  // Escape at the start of every line (`m` flag), not just the paragraph: an
  // inline break (soft break / page break) can put block syntax at the start of
  // a later line. Only horizontal whitespace `[ \t]` leads the marker so the
  // anchor doesn't cross the newline.
  return text
    .replace(/^(?<ws>[ \t]*)(?<marker>[#>=])/gmu, "$<ws>\\$<marker>")
    .replace(/^(?<ws>[ \t]*)(?<marker>[-+*])(?=[ \t]|$)/gmu, "$<ws>\\$<marker>")
    .replace(/^(?<ws>[ \t]*)-(?=[- \t]*$)/gmu, "$<ws>\\-")
    .replace(
      /^(?<ws>[ \t]*)(?<num>\d{1,9})(?<marker>[.)])(?<after>[ \t])/gmu,
      "$<ws>$<num>\\$<marker>$<after>",
    );
}

/**
 * Text that is not Markdown (a `text/plain` alternative-format chunk) as a
 * Markdown block that reads back as the same literal text.
 */
export const escapePlainTextBlock = (text: string): string =>
  escapeLeadingBlockMarker(escapeInline(text));
