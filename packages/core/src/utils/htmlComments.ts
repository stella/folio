/**
 * Index just past the HTML comment that opens at `start`, or `null` when the
 * comment is never closed.
 *
 * `<!-->` and `<!--->` are ABRUPT-CLOSING comments: the tokenizer's comment
 * start and comment-start-dash states end the comment at that `>` (HTML
 * §13.2.5.42-43). Scanning for `-->` from `start + 4` runs straight past them,
 * finds nothing, and every caller then treats the rest of the document as one
 * unterminated comment — a pasted payload that opens with `<!-->` is dropped
 * whole and silently.
 *
 * Both comment strippers resolve the terminator here so the two cannot drift.
 */
export const htmlCommentEnd = (html: string, start: number): number | null => {
  if (html[start + 4] === ">") {
    return start + 5;
  }
  if (html[start + 4] === "-" && html[start + 5] === ">") {
    return start + 6;
  }
  const end = html.indexOf("-->", start + 4);
  return end === -1 ? null : end + 3;
};
