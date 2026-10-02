/** Physical revision wrappers may split around hyperlinks without changing their meaning. */

const ADJACENT_REVISION_SPANS = /(<(ins|del)\b[^>]*>)((?:(?!<\/?(?:ins|del)\b)[\s\S])*)<\/\2>\1/gu;

/** Keep revision kinds, authors, dates, text and link targets; omit volatile IDs and wrapper splits. */
export const comparableMarkdown = (markdown: string): string => {
  let normalized = markdown
    .replace(/(<(?:ins|del)\b[^>]*?) id="[^"]*"/gu, "$1")
    .replaceAll("****", "");
  while (true) {
    const merged = normalized.replace(ADJACENT_REVISION_SPANS, "$1$3");
    if (merged === normalized) return normalized;
    normalized = merged;
  }
};
