/**
 * Inline markdown (emphasis, strikethrough, code spans, links, line breaks)
 * rendered into document runs. Shared by the plain markdown profile and the
 * legal-source compiler so both surfaces read `**bold**` the same way.
 */
import type { Token } from "marked";

import type { ParagraphContent, Run } from "../model/document";
import { sanitizeMarkdownHref } from "./href";
import { isTokenType, lexInlineMarkdown } from "./lexer";

// Whitelisted by folio-core's `toMarkdown` monospace inference, so a code span
// survives a DOCX → markdown round-trip. Folio renders Courier New through its
// bundled Cousine face.
const MONO_FONT = { ascii: "Courier New", hAnsi: "Courier New" } as const;

/** Run-level formatting an inline token adds to the text it wraps. */
export type InlineRunFormat = {
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  mono?: boolean;
  highlight?: "yellow";
};

export type InlineMarkdownOptions = {
  /** Formatting every run starts from (a bold clause heading, an italic recital). */
  base?: InlineRunFormat;
  /**
   * Wrap each `[[…]]` span in a yellow highlight. Legal drafts mark the
   * points a reviewing lawyer still has to fill in that way; plain markdown
   * keeps the brackets as text.
   */
  placeholders?: boolean;
};

/**
 * One run. A Word run cannot carry a raw newline (the layout engine renders
 * such lines on top of each other), so "\n" becomes an explicit break node.
 */
export const textRun = (text: string, format: InlineRunFormat = {}): Run => {
  const formatting = {
    ...(format.bold ? { bold: true } : {}),
    ...(format.italic ? { italic: true } : {}),
    ...(format.strike ? { strike: true } : {}),
    ...(format.mono ? { fontFamily: MONO_FONT } : {}),
    ...(format.highlight ? { highlight: format.highlight } : {}),
  };
  const content: Run["content"] = [];
  for (const [index, segment] of text.split("\n").entries()) {
    if (index > 0) {
      content.push({ type: "break" });
    }
    if (segment.length > 0) {
      content.push({ type: "text", text: segment, preserveSpace: true });
    }
  }
  if (content.length === 0) {
    content.push({ type: "text", text: "", preserveSpace: true });
  }
  return { type: "run", formatting, content };
};

// Split text on `[[…]]` markers into a sequence of runs. Each placeholder
// becomes its own highlighted run; the surrounding text inherits the base
// formatting but not the highlight.
const PLACEHOLDER_PATTERN = /\[\[(?<inner>[^\][]+?)\]\]/gu;

/**
 * Literal text (no markdown reading) with `[[…]]` placeholders highlighted:
 * for fields that are data rather than prose, such as signature parties.
 */
export const plainTextRuns = (text: string, format: InlineRunFormat = {}): Run[] =>
  placeholderRuns(text, format);

const placeholderRuns = (text: string, format: InlineRunFormat): Run[] => {
  if (!text.includes("[[")) {
    return [textRun(text, format)];
  }
  const runs: Run[] = [];
  let cursor = 0;
  for (const match of text.matchAll(PLACEHOLDER_PATTERN)) {
    const start = match.index;
    if (start > cursor) {
      runs.push(textRun(text.slice(cursor, start), format));
    }
    runs.push(textRun(match.groups?.["inner"] ?? "", { ...format, highlight: "yellow" }));
    cursor = start + match[0].length;
  }
  if (cursor < text.length) {
    runs.push(textRun(text.slice(cursor), format));
  }
  return runs.length > 0 ? runs : [textRun(text, format)];
};

const plainRuns = (text: string, format: InlineRunFormat, placeholders: boolean): Run[] =>
  placeholders ? placeholderRuns(text, format) : [textRun(text, format)];

type InlineContext = {
  placeholders: boolean;
};

const tokensToRuns = (
  tokens: Token[] | undefined,
  fallback: string,
  format: InlineRunFormat,
  context: InlineContext,
): ParagraphContent[] => {
  if (!tokens || tokens.length === 0) {
    return plainRuns(fallback, format, context.placeholders);
  }
  const runs: ParagraphContent[] = [];
  for (const token of tokens) {
    if (isTokenType(token, "strong")) {
      runs.push(...tokensToRuns(token.tokens, token.text, { ...format, bold: true }, context));
    } else if (isTokenType(token, "em")) {
      runs.push(...tokensToRuns(token.tokens, token.text, { ...format, italic: true }, context));
    } else if (isTokenType(token, "del")) {
      runs.push(...tokensToRuns(token.tokens, token.text, { ...format, strike: true }, context));
    } else if (isTokenType(token, "codespan")) {
      runs.push(textRun(token.text, { ...format, mono: true }));
    } else if (isTokenType(token, "link")) {
      runs.push(...linkRuns(token.tokens, token.text, token.href, format, context));
    } else if (isTokenType(token, "paragraph")) {
      runs.push(...tokensToRuns(token.tokens, token.text, format, context));
    } else if (token.type === "br") {
      runs.push({ type: "run", content: [{ type: "break" }] });
    } else if (token.type === "space") {
      if (runs.length > 0 && token.raw.includes("\n")) {
        runs.push(textRun("\n", format));
      }
    } else if (isTokenType(token, "text")) {
      const nested = token.tokens;
      if (nested && nested.length > 0) {
        runs.push(...tokensToRuns(nested, token.text, format, context));
      } else {
        runs.push(...plainRuns(token.text, format, context.placeholders));
      }
    } else if ("text" in token && typeof token.text === "string") {
      runs.push(...plainRuns(token.text, format, context.placeholders));
    }
  }
  return runs.length > 0 ? runs : plainRuns(fallback, format, context.placeholders);
};

const linkRuns = (
  tokens: Token[] | undefined,
  text: string,
  rawHref: string,
  format: InlineRunFormat,
  context: InlineContext,
): ParagraphContent[] => {
  const children = tokensToRuns(tokens, text, format, context).filter(
    (child): child is Run => child.type === "run",
  );
  const linkChildren = children.length > 0 ? children : [textRun(text, format)];
  const href = sanitizeMarkdownHref(rawHref);
  if (!href) {
    return linkChildren;
  }
  const anchor = href.startsWith("#") ? href.slice(1) : undefined;
  return [
    {
      type: "hyperlink",
      href,
      ...(anchor ? { anchor } : {}),
      children: linkChildren,
    },
  ];
};

/** Render already-lexed inline tokens; `fallback` is the source text when the tokens are empty. */
export const inlineTokensToRuns = (
  tokens: Token[] | undefined,
  fallback: string,
  options: InlineMarkdownOptions = {},
): ParagraphContent[] =>
  tokensToRuns(tokens, fallback, options.base ?? {}, {
    placeholders: options.placeholders ?? false,
  });

/** Lex and render one paragraph's inline markdown. */
export const inlineMarkdownToRuns = (
  text: string,
  options: InlineMarkdownOptions = {},
): ParagraphContent[] => inlineTokensToRuns(lexInlineMarkdown(text), text, options);
