/**
 * Remove complete `<?xml ...?>` declarations from attacker-controlled clipboard HTML.
 *
 * Linear scan instead of `/<\?xml[^>]*>/gi`: that pattern backtracks polynomially
 * when many `<?xml` openers appear with no closing `>`. An unterminated opener
 * is preserved verbatim (same as the regex, which never matched it).
 */
export function stripXmlDeclarations(html: string): string {
  const chunks: string[] = [];
  let cursor = 0;
  let searchFrom = 0;

  while (true) {
    const start = html.indexOf("<", searchFrom);
    if (start === -1) {
      chunks.push(html.slice(cursor));
      break;
    }

    const isXmlDeclaration =
      html.charCodeAt(start + 1) === 63 &&
      (html.charCodeAt(start + 2) === 88 || html.charCodeAt(start + 2) === 120) &&
      (html.charCodeAt(start + 3) === 77 || html.charCodeAt(start + 3) === 109) &&
      (html.charCodeAt(start + 4) === 76 || html.charCodeAt(start + 4) === 108);
    if (!isXmlDeclaration) {
      searchFrom = start + 1;
      continue;
    }

    const end = html.indexOf(">", start + 5);
    if (end === -1) {
      chunks.push(html.slice(cursor));
      break;
    }

    chunks.push(html.slice(cursor, start));
    cursor = end + 1;
    searchFrom = cursor;
  }

  return chunks.join("");
}
