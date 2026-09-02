/**
 * One `marked` configuration for every markdown surface in docx-core. Plain
 * markdown (`compileMarkdownToContent`) and the legal-source compiler read the
 * same GFM token stream; the legal profile adds a block-level extension that
 * turns an `@directive` line into its own token so the line never merges into
 * a neighbouring paragraph.
 */
import { Marked, type Token, type TokenizerExtension, type Tokens } from "marked";

/** A `@directive [argument]` line of a legal draft, tokenized as its own block. */
export type LegalDirectiveToken = {
  type: "legalDirective";
  raw: string;
  /** Lower-cased, with the leading `@` (`@clause`, `@doc`). */
  directive: string;
  /** The rest of the line, trimmed; empty when the directive stands alone. */
  argument: string;
};

const LEGAL_DIRECTIVE_TOKEN_TYPE = "legalDirective";

// A directive is `@` plus an identifier at the start of a line. Leading
// indentation is tolerated (models indent freely); `@` followed by anything
// else (`@ 5%`) is ordinary prose.
const DIRECTIVE_LINE_PATTERN =
  /^[ \t]*@(?<name>[A-Za-z][A-Za-z0-9_-]*)(?<argument>[^\n]*)(?:\n|$)/u;
// Line-anchored (`m`) so it serves both the block-start lookahead over a
// multi-line remainder and the per-line boundary pass in `lexLegalSource`.
const DIRECTIVE_LINE_START_PATTERN = /^[ \t]*@[A-Za-z]/mu;

const legalDirectiveExtension: TokenizerExtension = {
  name: LEGAL_DIRECTIVE_TOKEN_TYPE,
  level: "block",
  // Lets marked's paragraph tokenizer stop a paragraph right before the next
  // directive line instead of swallowing it as a soft-wrapped continuation.
  start: (src) => {
    const index = src.search(DIRECTIVE_LINE_START_PATTERN);
    return index === -1 ? undefined : index;
  },
  tokenizer: (src) => {
    const match = DIRECTIVE_LINE_PATTERN.exec(src);
    const name = match?.groups?.["name"];
    if (!match || name === undefined) {
      return undefined;
    }
    const token: LegalDirectiveToken = {
      type: LEGAL_DIRECTIVE_TOKEN_TYPE,
      raw: match[0],
      directive: `@${name.toLowerCase()}`,
      argument: (match.groups?.["argument"] ?? "").trim(),
    };
    return token;
  },
};

// Instances, not `marked.use`: the global `marked` singleton is shared with
// every other consumer in the process, and the directive extension must only
// apply to legal drafts.
const plainMarkdown = new Marked({ gfm: true });
const legalMarkdown = new Marked({ gfm: true, extensions: [legalDirectiveExtension] });

/** GFM block tokens of a markdown document. */
export const lexMarkdown = (source: string): Token[] => plainMarkdown.lexer(source);

export type LegalSourceLexResult = {
  tokens: Token[];
  /** Maps a line number in the lexed text back to the author's source. */
  originalLineOf: (line: number) => number;
};

/**
 * A directive line always starts a block. marked's paragraph tokenizer asks
 * block extensions where the next block starts, but its list and blockquote
 * tokenizers treat any following non-blank line as lazy continuation, so
 * `- item\n@clause Next` would swallow the directive. Inserting one blank
 * line before every directive that lacks one gives every tokenizer the same
 * boundary; the returned map keeps diagnostics on the author's line numbers.
 */
const separateDirectiveLines = (
  source: string,
): { text: string; originalLineOf: (line: number) => number } => {
  const lines = source.split("\n");
  const output: string[] = [];
  // insertedBefore[i]: blank lines added ahead of output line i (0-based).
  const insertedBefore: number[] = [];
  let inserted = 0;
  for (const [index, line] of lines.entries()) {
    const previous = output.at(-1);
    if (
      index > 0 &&
      DIRECTIVE_LINE_START_PATTERN.test(line) &&
      previous !== undefined &&
      previous.trim() !== ""
    ) {
      output.push("");
      insertedBefore.push(inserted);
      inserted += 1;
    }
    output.push(line);
    insertedBefore.push(inserted);
  }
  return {
    text: output.join("\n"),
    originalLineOf: (line) => line - (insertedBefore.at(line - 1) ?? inserted),
  };
};

/** GFM block tokens of a legal draft, with `@directive` lines as {@link LegalDirectiveToken}s. */
export const lexLegalSource = (source: string): LegalSourceLexResult => {
  const prepared = separateDirectiveLines(source);
  return {
    tokens: legalMarkdown.lexer(prepared.text),
    originalLineOf: prepared.originalLineOf,
  };
};

/** Inline tokens (emphasis, code spans, links, breaks) of one paragraph's text. */
export const lexInlineMarkdown = (text: string): Token[] =>
  plainMarkdown.Lexer.lexInline(text, plainMarkdown.defaults);

export const isLegalDirectiveToken = (token: Token): token is LegalDirectiveToken =>
  token.type === LEGAL_DIRECTIVE_TOKEN_TYPE;

// marked's `Token` union carries a `Tokens.Generic` member whose `type: string`
// overlaps every literal (and whose `any` index signature absorbs the whole
// union, so `Exclude` can't strip it). A plain `token.type === "x"` guard thus
// narrows to `Tokens.X | Tokens.Generic` and field access stays `any`. This
// predicate maps the discriminator to the concrete token; the runtime type is
// exactly the one the discriminator matched because marked only emits Generic
// for custom extensions, and the one extension registered here has its own
// predicate above.
type HandledTokens = {
  blockquote: Tokens.Blockquote;
  code: Tokens.Code;
  codespan: Tokens.Codespan;
  del: Tokens.Del;
  em: Tokens.Em;
  heading: Tokens.Heading;
  html: Tokens.HTML;
  link: Tokens.Link;
  list: Tokens.List;
  paragraph: Tokens.Paragraph;
  strong: Tokens.Strong;
  table: Tokens.Table;
  text: Tokens.Text;
};

export const isTokenType = <T extends keyof HandledTokens>(
  token: Token,
  type: T,
): token is HandledTokens[T] => token.type === type;
