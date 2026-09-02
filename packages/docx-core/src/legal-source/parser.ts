/**
 * Legal-source parser: GFM markdown plus `@` directives, read from one
 * `marked` token stream (see `../markdown/lexer.ts`). Directive lines
 * structure the draft (title, recitals, clauses, schedules, signatures); the
 * markdown between them is ordinary GFM, so headings, emphasis, links, lists,
 * and pipe tables all mean what they mean in markdown. Paragraph strings on
 * the draft keep their inline markdown; the compiler renders it into runs.
 *
 * Three directive bodies stay line-oriented on purpose: `@list` (manual
 * markers such as `2)` or `3.1` are stripped rather than parsed), `@table`
 * (a pipe table without the `---` separator row still counts), and
 * `@signatures` (`key: value` lines). Their bodies are taken as raw text.
 */
import type { Token, Tokens } from "marked";

import { isLegalDirectiveToken, isTokenType, lexLegalSource } from "../markdown/lexer";
import type { LegalDirectiveToken, LegalSourceLexResult } from "../markdown/lexer";
import type {
  Autofix,
  LegalDraft,
  LegalDraftBlock,
  LegalDraftDiagnostic,
  LegalDocumentKind,
  LegalNumberingProfile,
  LegalPageOrientation,
  LegalPageSize,
  LegalSignatureParty,
  LegalSourceParseResult,
} from "./types";

const DEFAULT_KIND = "agreement" satisfies LegalDocumentKind;
const DEFAULT_LOCALE = "en-GB";
const DEFAULT_NUMBERING = "legal" satisfies LegalNumberingProfile;
const DEFAULT_PAGE_SIZE = "A4" satisfies LegalPageSize;
const DEFAULT_ORIENTATION = "portrait" satisfies LegalPageOrientation;
const MAX_CLAUSE_LEVEL = 6;

const DIRECTIVE_ALIASES: Record<string, string> = {
  "@annex": "@schedule",
  "@appendix": "@schedule",
  "@body": "@paragraph",
  "@para": "@paragraph",
  // @preamble was a thin styling variant of @paragraph; alias so any
  // legacy source still parses.
  "@preamble": "@paragraph",
  "@section": "@clause",
  "@signature": "@signatures",
  "@subsection": "@subclause",
};

const CLOSING_DIRECTIVE_PATTERN = /^@end[a-z]*$/u;

const DIRECTIVES = new Set([
  "@doc",
  "@title",
  "@recital",
  "@clause",
  "@subclause",
  "@paragraph",
  "@list",
  "@table",
  "@schedule",
  "@signatures",
  "@pagebreak",
]);

/**
 * Walks the top-level token stream. Line numbers come from the newlines in
 * the consumed tokens' `raw` text, which marked keeps contiguous with the
 * source, so diagnostics point at the directive line that produced them.
 */
class TokenCursor {
  private readonly tokens: readonly Token[];
  private readonly originalLineOf: (line: number) => number;
  private index = 0;
  private lineNumber = 1;

  constructor({ tokens, originalLineOf }: LegalSourceLexResult) {
    this.tokens = tokens;
    this.originalLineOf = originalLineOf;
  }

  get current(): Token | undefined {
    return this.tokens.at(this.index);
  }

  /** The author's line number for the current token. */
  get line(): number {
    return this.originalLineOf(this.lineNumber);
  }

  /** Index of the current token; lets a caller tell whether a helper consumed anything. */
  get position(): number {
    return this.index;
  }

  advance(): Token | undefined {
    const token = this.current;
    if (token === undefined) {
      return undefined;
    }
    this.index += 1;
    this.lineNumber += countNewlines(token.raw);
    return token;
  }
}

const countNewlines = (text: string): number => {
  let count = 0;
  for (const char of text) {
    if (char === "\n") {
      count += 1;
    }
  }
  return count;
};

/** A token that starts a new structural block, so no body may run past it. */
const startsStructure = (token: Token): boolean =>
  isLegalDirectiveToken(token) || isTokenType(token, "heading");

type ParseState = {
  blocks: LegalDraftBlock[];
  diagnostics: LegalDraftDiagnostic[];
  fixes: Autofix[];
  meta: LegalDraft["meta"];
  cursor: TokenCursor;
};

export const parseLegalSource = (
  source: string,
  options: { titleFallback?: string } = {},
): LegalSourceParseResult => {
  const state: ParseState = {
    blocks: [],
    diagnostics: [],
    fixes: [],
    meta: {
      kind: DEFAULT_KIND,
      locale: DEFAULT_LOCALE,
      numbering: DEFAULT_NUMBERING,
      page: {
        size: DEFAULT_PAGE_SIZE,
        orientation: DEFAULT_ORIENTATION,
      },
      title: null,
    },
    cursor: new TokenCursor(lexLegalSource(source)),
  };

  for (;;) {
    const token = state.cursor.current;
    if (token === undefined) {
      break;
    }
    const line = state.cursor.line;
    if (isLegalDirectiveToken(token)) {
      state.cursor.advance();
      parseDirective(token, line, state);
      continue;
    }
    if (isTokenType(token, "heading")) {
      state.cursor.advance();
      parseMarkdownHeading(token, line, state);
      continue;
    }
    parseBareMarkdown(state);
  }

  if (!state.meta.title) {
    const firstTitle = state.blocks.find((block) => block.type === "title");
    state.meta.title =
      firstTitle?.type === "title"
        ? firstTitle.text
        : (options.titleFallback ?? "Untitled document");
  }

  const draft: LegalDraft = { meta: state.meta, blocks: state.blocks };
  return applyDocumentAutofixes({ diagnostics: state.diagnostics, draft, fixes: state.fixes });
};

const pushBlock = (state: ParseState, block: LegalDraftBlock | null) => {
  if (!block) {
    return;
  }
  state.blocks.push(block);
  if (block.type === "title") {
    state.meta.title = block.text;
  }
};

const parseDirective = (token: LegalDirectiveToken, line: number, state: ParseState) => {
  const { diagnostics, fixes, meta } = state;
  const rawDirective = token.directive;
  const directive = DIRECTIVE_ALIASES[rawDirective] ?? rawDirective;
  const { argument } = token;

  // Models raised on HTML and LaTeX close what they open (`@endlist`,
  // `@end`). Blocks here end where the next one starts, so a closer changes
  // nothing; accept it and say so rather than failing the whole draft.
  if (CLOSING_DIRECTIVE_PATTERN.test(directive)) {
    fixes.push({
      code: "closing-directive-ignored",
      message: `Ignored ${rawDirective}: blocks end where the next directive starts.`,
      line,
    });
    return;
  }

  if (!DIRECTIVES.has(directive)) {
    // Report the directive as the author spelled it, not lower-cased.
    const spelled = token.raw.trim().split(/\s+/u).at(0) ?? rawDirective;
    diagnostics.push({
      code: "unknown-directive",
      message: `Unknown legal directive "${spelled}".`,
      severity: "error",
      line,
    });
    return;
  }

  if (directive !== rawDirective) {
    fixes.push({
      code: "directive-alias-normalized",
      message: `Normalized ${rawDirective} to ${directive}.`,
      line,
    });
  }

  switch (directive) {
    case "@doc":
      parseDocDirective(argument, meta, diagnostics, line);
      return;
    case "@title":
      // Title blocks have no body: the argument is the title, and the next
      // block starts fresh (a bare paragraph after `@title` is body text).
      pushBlock(state, argument ? { type: "title", text: argument } : null);
      return;
    case "@recital":
      pushBlock(state, { type: "recital", paragraphs: takeParagraphs(state) });
      return;
    case "@clause":
      pushBlock(state, clauseBlock(1, argument, line, state));
      return;
    case "@subclause":
      pushBlock(state, clauseBlock(2, argument, line, state));
      return;
    case "@paragraph":
      pushBlock(state, { type: "paragraph", paragraphs: takeParagraphs(state) });
      return;
    case "@list": {
      const ordered = /\bordered\b/iu.test(argument);
      const items = takeRawLines(state).flatMap((rawLine) => {
        const stripped = stripListMarker(rawLine, ordered);
        return stripped ? [stripped] : [];
      });
      pushBlock(state, { type: "list", ordered, items });
      return;
    }
    case "@table":
      pushBlock(state, parseTableBlock(takeRawLines(state), line, diagnostics, fixes));
      return;
    case "@schedule":
      pushBlock(state, {
        type: "schedule",
        heading: stripManualNumbering(argument, line, fixes),
        paragraphs: takeParagraphs(state),
      });
      return;
    case "@signatures":
      pushBlock(state, {
        type: "signatures",
        parties: parseSignatureParties(takeRawLines(state), argument),
      });
      return;
    case "@pagebreak":
      pushBlock(state, { type: "pageBreak" });
      return;
    default:
      return;
  }
};

const clauseBlock = (
  level: number,
  rawHeading: string,
  line: number,
  state: ParseState,
): LegalDraftBlock => {
  const heading = stripManualNumbering(rawHeading, line, state.fixes);
  const paragraphs = takeParagraphs(state);
  // The AI sometimes uses `@clause` as a generic "section" wrapper without
  // giving it a title. Rather than rejecting, downgrade to a plain paragraph
  // block: same body content, no clause numbering or heading row. The fix
  // log keeps the event visible without blocking compile.
  if (!heading) {
    state.fixes.push({
      code: "headingless-clause-downgraded",
      message: "Converted a headingless @clause into a paragraph block.",
      line,
    });
    return { type: "paragraph", paragraphs };
  }
  return { type: "clause", level, heading, paragraphs };
};

// A markdown heading is the directive-free spelling of a title (`#`) or a
// clause (`##` and deeper), so a model that writes plain markdown still gets
// legal structure.
const parseMarkdownHeading = (token: Tokens.Heading, line: number, state: ParseState) => {
  const heading = token.text.trim();
  state.fixes.push({
    code: "markdown-heading-normalized",
    message: "Converted a Markdown heading into a legal directive.",
    line,
  });
  if (token.depth === 1) {
    pushBlock(state, heading ? { type: "title", text: heading } : null);
    return;
  }
  pushBlock(state, clauseBlock(Math.min(token.depth - 1, MAX_CLAUSE_LEVEL), heading, line, state));
};

/**
 * Markdown outside any directive. Consecutive prose stays one paragraph
 * block (several paragraphs), while a list or table is its own block so
 * markdown lists get real numbering instead of literal `-` markers.
 */
const parseBareMarkdown = (state: ParseState) => {
  const token = state.cursor.current;
  if (token === undefined) {
    return;
  }
  if (isTokenType(token, "list")) {
    state.cursor.advance();
    pushBlock(state, {
      type: "list",
      ordered: token.ordered,
      items: flattenListItems(token),
    });
    return;
  }
  if (isTokenType(token, "table")) {
    state.cursor.advance();
    pushBlock(state, {
      type: "table",
      table: {
        headers: token.header.map((cell) => cellText(cell)),
        rows: token.rows.map((row) => row.map((cell) => cellText(cell))),
      },
    });
    return;
  }
  const before = state.cursor.position;
  const paragraphs = takeParagraphs(state);
  if (paragraphs.length > 0) {
    pushBlock(state, { type: "paragraph", paragraphs });
    return;
  }
  // Nothing prose-like here. When the helper already consumed whitespace or
  // a rule, the cursor now sits on a directive or heading that the main loop
  // must dispatch; only a token the helper refused to touch is skipped.
  if (state.cursor.position === before) {
    state.cursor.advance();
  }
};

/**
 * Consecutive prose tokens as paragraph strings: paragraphs (soft line
 * breaks collapsed to spaces), blockquotes, code blocks, and raw HTML. Stops
 * at the next directive, heading, list, or table so those keep their order.
 */
const takeParagraphs = (state: ParseState): string[] => {
  const paragraphs: string[] = [];
  for (;;) {
    const token = state.cursor.current;
    if (token === undefined || startsStructure(token)) {
      return paragraphs;
    }
    if (isTokenType(token, "list") || isTokenType(token, "table")) {
      return paragraphs;
    }
    const prose = proseParagraphs(token);
    if (prose === null) {
      return paragraphs;
    }
    state.cursor.advance();
    paragraphs.push(...prose);
  }
};

/** The paragraph strings of one prose token, or `null` when the token is not prose. */
const proseParagraphs = (token: Token): string[] | null => {
  if (token.type === "space" || token.type === "hr") {
    return [];
  }
  if (isTokenType(token, "paragraph") || isTokenType(token, "text")) {
    const text = collapseSoftBreaks(token.text);
    return text ? [text] : [];
  }
  if (isTokenType(token, "blockquote")) {
    return token.tokens.flatMap((inner) => proseParagraphs(inner) ?? []);
  }
  if (isTokenType(token, "code")) {
    // Fenced code is literal by markdown semantics; the compiler reads every
    // paragraph string as inline markdown, so escape the characters that
    // would otherwise turn into emphasis, links, or placeholders.
    return token.text.split("\n").flatMap((codeLine) => {
      const trimmed = codeLine.trim();
      return trimmed ? [escapeInlineMarkdown(trimmed)] : [];
    });
  }
  if (isTokenType(token, "html")) {
    const text = collapseSoftBreaks(token.text);
    return text ? [text] : [];
  }
  return null;
};

/**
 * Raw source lines of everything up to the next directive or heading, for
 * the line-oriented directive bodies (`@list`, `@table`, `@signatures`).
 */
const takeRawLines = (state: ParseState): string[] => {
  let raw = "";
  for (;;) {
    const token = state.cursor.current;
    if (token === undefined || startsStructure(token)) {
      break;
    }
    state.cursor.advance();
    raw += token.raw;
  }
  return raw.split("\n").map((rawLine) => rawLine.trimEnd());
};

// Nested items lose their depth (the draft's list block is flat), matching
// how indented `- ` lines under `@list` have always read.
const flattenListItems = (list: Tokens.List): string[] => {
  const items: string[] = [];
  for (const item of list.items) {
    const parts: string[] = [];
    const nested: Tokens.List[] = [];
    for (const child of item.tokens) {
      if (isTokenType(child, "list")) {
        nested.push(child);
        continue;
      }
      parts.push(...(proseParagraphs(child) ?? []));
    }
    const text = parts.join(" ").trim();
    if (text) {
      items.push(text);
    }
    for (const nestedList of nested) {
      items.push(...flattenListItems(nestedList));
    }
  }
  return items;
};

const cellText = (cell: Tokens.TableCell): string => collapseSoftBreaks(cell.text);

const INLINE_MARKDOWN_SPECIALS = /[\\`*_[\]<>~!]/gu;

/** Backslash-escape the inline markdown syntax so the text renders verbatim. */
const escapeInlineMarkdown = (text: string): string =>
  text.replaceAll(INLINE_MARKDOWN_SPECIALS, (char) => `\\${char}`);

// A paragraph's soft-wrapped lines are one paragraph; trim and join with a
// single space so the wrapped markdown reads as continuous prose.
const collapseSoftBreaks = (text: string): string =>
  text
    .split("\n")
    .map((textLine) => textLine.trim())
    .filter((textLine) => textLine.length > 0)
    .join(" ");

const parseDocDirective = (
  argument: string,
  meta: LegalDraft["meta"],
  diagnostics: LegalDraftDiagnostic[],
  line: number,
) => {
  const attrs = parseAttributes(argument);

  const kind = attrs.get("kind");
  if (kind !== undefined && isLegalKind(kind)) {
    meta.kind = kind;
  } else if (kind !== undefined) {
    diagnostics.push({
      code: "invalid-doc-attribute",
      message: `Invalid @doc kind "${kind}".`,
      severity: "warning",
      line,
    });
  }

  const locale = attrs.get("locale");
  if (locale) {
    meta.locale = locale;
  }

  const numbering = attrs.get("numbering");
  if (numbering !== undefined && isNumberingProfile(numbering)) {
    meta.numbering = numbering;
  } else if (numbering !== undefined) {
    diagnostics.push({
      code: "invalid-doc-attribute",
      message: `Invalid @doc numbering "${numbering}".`,
      severity: "warning",
      line,
    });
  }

  const page = attrs.get("page");
  if (page !== undefined && isPageSize(page)) {
    meta.page.size = page;
  } else if (page !== undefined) {
    diagnostics.push({
      code: "invalid-doc-attribute",
      message: `Invalid @doc page "${page}".`,
      severity: "warning",
      line,
    });
  }

  const orientation = attrs.get("orientation");
  if (orientation !== undefined && isPageOrientation(orientation)) {
    meta.page.orientation = orientation;
  } else if (orientation !== undefined) {
    diagnostics.push({
      code: "invalid-doc-attribute",
      message: `Invalid @doc orientation "${orientation}".`,
      severity: "warning",
      line,
    });
  }

  const title = attrs.get("title");
  if (title) {
    meta.title = title;
  }

  for (const key of attrs.keys()) {
    if (!["kind", "locale", "numbering", "page", "orientation", "title"].includes(key)) {
      diagnostics.push({
        code: "unknown-doc-attribute",
        message: `Unknown @doc attribute "${key}".`,
        severity: "warning",
        line,
      });
    }
  }
};

const parseAttributes = (value: string): Map<string, string> => {
  const attrs = new Map<string, string>();
  let index = 0;

  while (index < value.length) {
    const attr = readAttribute(value, index);
    if (!attr) {
      index = skipMalformedAttribute(value, index);
      continue;
    }

    attrs.set(attr.key, attr.value);
    index = attr.nextIndex;
  }
  return attrs;
};

type ParsedAttribute = {
  key: string;
  value: string;
  nextIndex: number;
};

const readAttribute = (value: string, startIndex: number): ParsedAttribute | null => {
  const keyStart = skipWhitespace(value, startIndex);
  const keyEnd = readAttributeKeyEnd(value, keyStart);
  const key = value.slice(keyStart, keyEnd).toLowerCase();
  const equalsIndex = skipWhitespace(value, keyEnd);
  if (!key || value.charAt(equalsIndex) !== "=") {
    return null;
  }

  const attributeValue = readAttributeValue(value, skipWhitespace(value, equalsIndex + 1));
  return {
    key,
    value: attributeValue.value,
    nextIndex: attributeValue.nextIndex,
  };
};

const skipWhitespace = (value: string, startIndex: number): number => {
  let index = startIndex;
  while (index < value.length && isWhitespace(value.charAt(index))) {
    index++;
  }
  return index;
};

const readAttributeKeyEnd = (value: string, startIndex: number): number => {
  let index = startIndex;
  while (index < value.length && isAttributeKeyChar(value.charAt(index))) {
    index++;
  }
  return index;
};

const readAttributeValue = (
  value: string,
  startIndex: number,
): { value: string; nextIndex: number } => {
  const quote = value.charAt(startIndex);
  if (quote === '"' || quote === "'") {
    return readQuotedAttributeValue(value, startIndex + 1, quote);
  }

  let index = startIndex;
  while (index < value.length && !isWhitespace(value.charAt(index))) {
    index++;
  }
  return { value: value.slice(startIndex, index), nextIndex: index };
};

const readQuotedAttributeValue = (
  value: string,
  startIndex: number,
  quote: string,
): { value: string; nextIndex: number } => {
  let index = startIndex;
  while (index < value.length && value.charAt(index) !== quote) {
    index++;
  }
  return {
    value: value.slice(startIndex, index),
    nextIndex: index < value.length ? index + 1 : index,
  };
};

const skipMalformedAttribute = (value: string, startIndex: number): number => {
  let index = skipWhitespace(value, startIndex);
  while (index < value.length && !isWhitespace(value.charAt(index))) {
    index++;
  }
  return index === startIndex ? index + 1 : index;
};

const isWhitespace = (char: string): boolean =>
  char === " " || char === "\t" || char === "\r" || char === "\n";

const isAttributeKeyChar = (char: string): boolean =>
  isAsciiAlphaNumeric(char) || char === "_" || char === "." || char === "-";

const isAsciiAlphaNumeric = (char: string): boolean => {
  const code = char.codePointAt(0);
  if (code === undefined) {
    return false;
  }
  return (code >= 48 && code <= 57) || (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
};

const parseTableBlock = (
  lines: string[],
  line: number,
  diagnostics: LegalDraftDiagnostic[],
  fixes: Autofix[],
): LegalDraftBlock => {
  const tableLines = lines.flatMap((rawLine) => {
    const trimmed = rawLine.trim();
    // A leading pipe is enough to recognise a row; tolerate a missing trailing
    // pipe (a common hand-written/AI variant) so the row is not silently dropped.
    return trimmed.startsWith("|") ? [trimmed] : [];
  });
  const rows = tableLines.flatMap((tableLine) => {
    const row = parsePipeRow(tableLine);
    return row.length > 0 ? [row] : [];
  });

  const header = rows.at(0) ?? [];
  const bodyRows = rows.slice(1).filter((row) => !isMarkdownDividerRow(row));
  const normalizedRows = bodyRows.map((row) => {
    if (row.length === header.length) {
      return row;
    }
    fixes.push({
      code: "table-row-width-normalized",
      message: "Normalized a table row to match the header width.",
      line,
    });
    return header.map((_, index) => row.at(index) ?? "");
  });

  if (header.length === 0) {
    diagnostics.push({
      code: "missing-table-header",
      message: "Table directives must include a pipe-table header row.",
      severity: "error",
      line,
    });
  }

  return {
    type: "table",
    table: {
      headers: header,
      rows: normalizedRows,
    },
  };
};

const applyDocumentAutofixes = ({
  diagnostics,
  draft,
  fixes,
}: {
  diagnostics: LegalDraftDiagnostic[];
  draft: LegalDraft;
  fixes: Autofix[];
}): LegalSourceParseResult => {
  const blocks: LegalDraftBlock[] = [];
  const normalizedTitle = normalizeTitle(draft.meta.title ?? "");
  let hasSeenSubstantiveBlock = false;

  for (const block of draft.blocks) {
    if (
      !hasSeenSubstantiveBlock &&
      block.type === "clause" &&
      block.level === 1 &&
      normalizedTitle.length > 0 &&
      normalizeTitle(block.heading) === normalizedTitle
    ) {
      fixes.push({
        code: "duplicate-title-clause-removed",
        message: "Removed a first clause that duplicated the document title.",
      });
      continue;
    }
    blocks.push(block);
    hasSeenSubstantiveBlock ||= isSubstantiveDraftBlock(block);
  }

  const signatureIndex = blocks.findIndex((block) => block.type === "signatures");
  if (signatureIndex !== -1 && signatureIndex !== blocks.length - 1) {
    const [signatureBlock] = blocks.splice(signatureIndex, 1);
    if (signatureBlock) {
      blocks.push(signatureBlock);
      fixes.push({
        code: "signatures-moved-to-end",
        message: "Moved the signatures block to the end of the document.",
      });
    }
  }

  return {
    draft: { ...draft, blocks },
    diagnostics,
    fixes,
  };
};

const isSubstantiveDraftBlock = (block: LegalDraftBlock): boolean =>
  block.type !== "title" && block.type !== "pageBreak";

// Localized aliases for `@signatures` field keys. Lets the AI
// write `strana:` / `funkce:` etc. when the document is in Czech
// without forcing a separate parser per language. Add new aliases
// here as locales come online — the canonical (English) keys are
// what the rest of the parser branches on.
const SIGNATURE_KEY_ALIASES: Record<string, string> = {
  // Canonical
  party: "party",
  by: "by",
  name: "name",
  title: "title",
  date: "date",
  // Czech / Slovak
  strana: "party",
  podepisuje: "by",
  podpisuje: "by",
  jméno: "name",
  jmeno: "name",
  meno: "name",
  funkce: "title",
  funkcia: "title",
  datum: "date",
  // German
  partei: "party",
  unterzeichnet: "by",
  unterschreibt: "by",
  funktion: "title",
  // French
  partie: "party",
  signataire: "by",
  nom: "name",
  fonction: "title",
  // Spanish
  parte: "party",
  firmante: "by",
  nombre: "name",
  cargo: "title",
  fecha: "date",
  // Italian
  firmatario: "by",
  nome: "name",
  carica: "title",
  data: "date",
  // Polish
  imie: "name",
  imię: "name",
  stanowisko: "title",
  // Portuguese
  assinante: "by",
  // Dutch
  ondertekent: "by",
  naam: "name",
  functie: "title",
  // Hungarian
  fél: "party",
  fel: "party",
  aláírja: "by",
  alairja: "by",
  név: "name",
  nev: "name",
  beosztás: "title",
  beosztas: "title",
  dátum: "date",
};
const SIGNATURE_FIELD_KEYS = new Set(["party", "by", "name", "title", "date"]);

const parseSignatureParties = (lines: string[], heading: string): LegalSignatureParty[] => {
  const parties: LegalSignatureParty[] = [];
  let current: LegalSignatureParty | null = null;

  const pushCurrent = () => {
    if (!current?.name.trim()) {
      current = null;
      return;
    }
    parties.push({
      name: current.name.trim(),
      ...(current.signatory ? { signatory: current.signatory.trim() } : {}),
      ...(current.title ? { title: current.title.trim() } : {}),
    });
    current = null;
  };

  const startParty = (name: string) => {
    pushCurrent();
    current = { name };
  };

  if (heading.trim()) {
    startParty(heading.trim().replace(/^party:\s*/iu, ""));
  }

  for (const rawLine of lines) {
    const trimmed = rawLine.trim();
    if (!trimmed) {
      continue;
    }

    const separatorIndex = trimmed.indexOf(":");
    if (separatorIndex === -1) {
      startParty(trimmed);
      continue;
    }

    const rawKey = trimmed.slice(0, separatorIndex).toLowerCase();
    const canonicalKey = SIGNATURE_KEY_ALIASES[rawKey];
    if (!canonicalKey || !SIGNATURE_FIELD_KEYS.has(canonicalKey)) {
      startParty(trimmed);
      continue;
    }

    const value = trimmed.slice(separatorIndex + 1).trim();
    if (canonicalKey === "party") {
      startParty(value);
      continue;
    }
    current ??= { name: "" };
    if ((canonicalKey === "by" || canonicalKey === "name") && value) {
      current.signatory = value;
    }
    if (canonicalKey === "title" && value) {
      current.title = value;
    }
  }

  pushCurrent();

  const seen = new Set<string>();
  return parties.filter((party) => {
    const key = `${party.name}\u0000${party.signatory ?? ""}\u0000${party.title ?? ""}`;
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
};

const ORDERED_LIST_MARKER_RE = /^(?:\d+(?:\.\d+)+|\d+[.)])\s+/u;
const MANUAL_NUMBERING_PREFIX_RE = /^(?:\d+(?:\.\d+)+|\d+[.)]|[A-Za-z][.)]|\([a-zivx]+\))\s+/u;

const stripListMarker = (rawLine: string, ordered: boolean): string => {
  const trimmed = rawLine.trim();
  if (ordered) {
    return trimmed.replace(ORDERED_LIST_MARKER_RE, "");
  }
  return trimmed.replace(/^[-*•]\s+/u, "");
};

const parsePipeRow = (tableLine: string): string[] =>
  tableLine
    .replace(/^\|/u, "")
    .replace(/\|$/u, "")
    .split("|")
    .map((cell) => cell.trim());

const isMarkdownDividerRow = (row: string[]): boolean =>
  row.every((cell) => /^:?-{3,}:?$/u.test(cell));

const stripManualNumbering = (value: string, line: number, fixes: Autofix[]): string => {
  // Numeric and letter branches both require a delimiter so legit
  // headings starting with a year or single-word capital
  // ("2024 Compliance Obligations", "A Party's Obligations") are
  // not silently rewritten:
  //   - `\d+(?:\.\d+)+` accepts multi-level numbers (1.1, 1.1.1)
  //     where the dot itself is the delimiter.
  //   - `\d+[.)]` accepts a single number followed by '.' or ')'.
  //   - `[A-Za-z][.)]` accepts a letter followed by '.' or ')'.
  //   - `\([a-zivx]+\)` accepts parenthesised letters/roman.
  const stripped = value.trim().replace(MANUAL_NUMBERING_PREFIX_RE, "");
  if (stripped !== value.trim()) {
    fixes.push({
      code: "manual-numbering-stripped",
      message: "Removed manual numbering from a directive heading.",
      line,
    });
  }
  return stripped;
};

const normalizeTitle = (value: string): string =>
  value
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim();

const isLegalKind = (value: string | undefined): value is LegalDocumentKind =>
  value === "agreement" ||
  value === "letter" ||
  value === "memo" ||
  value === "checklist" ||
  value === "pleading" ||
  value === "other";

const isNumberingProfile = (value: string | undefined): value is LegalNumberingProfile =>
  value === "legal" || value === "none" || value === "checklist";

const isPageSize = (value: string | undefined): value is LegalPageSize =>
  value === "A4" || value === "Letter";

const isPageOrientation = (value: string | undefined): value is LegalPageOrientation =>
  value === "portrait" || value === "landscape";
