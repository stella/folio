/**
 * Group corpus failures into signatures a human can act on.
 *
 * Thousands of files produce a handful of distinct defects, each reported with
 * a different file name, element index and identifier. A signature is the
 * invariant that broke, the failure message with every per-file particular
 * erased, and the innermost folio frame that produced it. Two files with the
 * same signature are the same bug, so the census counts signatures and the
 * baseline ratchets on them.
 *
 * Signatures carry no line numbers: a signature that moved because an unrelated
 * edit shifted a file would read as a new defect and a disappeared one.
 */

import type { ExtendedCorpusInvariant } from "./corpus-invariants/contract";

export const CORPUS_INVARIANTS = {
  /** The worker returned a verdict at all: no hang, no process abort. */
  completes: "completes",
  parse: "parse",
  fixedPoint: "fixed-point",
  repackValidates: "repack-validates",
  styleSetRebuild: "style-set-rebuild",
} as const;

export type CorpusInvariant =
  | (typeof CORPUS_INVARIANTS)[keyof typeof CORPUS_INVARIANTS]
  | ExtendedCorpusInvariant;

export type CorpusFailure = {
  invariant: CorpusInvariant;
  message: string;
  frame: string;
};

/** The frame slot for a failure that is an assertion, not a thrown error. */
export const NO_FRAME = "-";

/**
 * How long a signature's message may be.
 *
 * Raised from 160 when path segments started naming the carrier they step
 * into: `content[paragraph].content[run]` says twice as much as `content[]`
 * and is twice as long. Over the committed baselines, 160 cut 200 of the 641
 * rows that carry a model path and 240 cuts 15, which is where the cap sat
 * before the carriers were named.
 */
export const MAX_MESSAGE_LENGTH = 240;

/**
 * The model path a difference message carries, from `package` to its leaf.
 *
 * One pattern, because two would drift: the normaliser shortens the path this
 * matches and `corpus-dispositions.ts` matches a disposition's pattern against
 * the path this finds. A segment holds no whitespace, colon or dot, so a long
 * path cannot swallow the rest of the message.
 */
export const MODEL_PATH_RE = /(?<![^\s])package(?:\.[^\s:.]+)*/u;

/** How a shortened path marks the segments it dropped, and a cut message its tail. */
export const ELISION = "…";

/**
 * The `type` discriminators a path segment may name.
 *
 * Closed on purpose. A path segment is the one place in a signature where a
 * value read out of the document could otherwise be spelled verbatim: `type`
 * is an ordinary key, and a package folio did not write can carry any string
 * under it. Only a discriminator the model itself declares is named; anything
 * else leaves the segment untyped, so the worst a package can do is make its
 * own losses harder to tell apart.
 *
 * `scripts/corpus-model-discriminators.test.ts` derives this set from
 * `packages/docx-core/src/model` and fails when the two disagree, so a member
 * added to the model cannot quietly stop naming its carrier.
 */
export const MODEL_TYPE_DISCRIMINATORS: ReadonlySet<string> = new Set([
  "band1Horz",
  "band1Vert",
  "band2Horz",
  "band2Vert",
  "blockSdt",
  "bookmarkEnd",
  "bookmarkStart",
  "break",
  "commentRangeEnd",
  "commentRangeStart",
  "commentReference",
  "complexField",
  "deletion",
  "drawing",
  "endnote",
  "endnoteRef",
  "fieldChar",
  "firstCol",
  "firstRow",
  "footer",
  "footnote",
  "footnoteRef",
  "header",
  "hyperlink",
  "inlineSdt",
  "inlineWrapper",
  "insertion",
  "instrText",
  "lastCol",
  "lastRow",
  "mathEquation",
  "moveFrom",
  "moveFromRangeEnd",
  "moveFromRangeStart",
  "moveTo",
  "moveToRangeEnd",
  "moveToRangeStart",
  "neCell",
  "noBreakHyphen",
  "nwCell",
  "paragraph",
  "paragraphPropertyChange",
  "preservedBlock",
  "preservedInline",
  "preservedXml",
  "renderedPageBreak",
  "run",
  "runPropertyChange",
  "seCell",
  "sectionPropertyChange",
  "shape",
  "simpleField",
  "softHyphen",
  "swCell",
  "symbol",
  "tab",
  "table",
  "tableCell",
  "tableCellPropertyChange",
  "tablePropertyChange",
  "tablePropertyExceptionChange",
  "tableRow",
  "tableRowPropertyChange",
  "text",
  "wholeTable",
]);

const GUID_RE = /\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/giu;
/**
 * An absolute path, and only an absolute path.
 *
 * The lookbehind is load-bearing. Without it the pattern matches from the first
 * separator of any relative path with three or more segments, so
 * `word/media/imageN.png` and `word/theme/themeN.xml` both collapse to
 * `word<path>` and two unrelated defects share one signature. A relative path
 * is the same on every machine, so there is nothing to erase in it.
 */
const ABSOLUTE_PATH_RE =
  /(?<![\w.@-])(?:file:\/\/)?(?:[A-Za-z]:)?[/\\](?:[\w.@-]+[/\\])+[\w.@-]+/gu;
const HEX_RUN_RE = /\b(?:0x)?[0-9a-f]{8,}\b/giu;
const DIGIT_RUN_RE = /\d+/gu;
const WHITESPACE_RE = /\s+/gu;
const STACK_FRAME_RE = /at (?:(?<fn>[^\s(]+) \()?(?<file>[^\s()]+?):\d+:\d+\)?/u;

/**
 * A message shortened by dropping the middle of the model path it carries.
 *
 * `undefined` when there is no path to shorten, or when not even the leaf fits
 * in what the prose around it leaves over.
 *
 * Which segment is worth keeping is not a matter of taste. A difference is
 * owned by the carrier it happened under and named by the field it happened
 * to, and both sit at the end: `…content[run].preservedAttributes` is the
 * defect, while `package.document.sections[].content[paragraph]` is where the
 * corpus happened to find it. Cutting from the right, which is what a plain
 * length cap does, throws away the only two segments that name the defect, and
 * a disposition pattern anchored on its leaf then stops claiming its own rows.
 * Nesting is unbounded — a table inside a cell inside a table — so no cap
 * makes this case go away.
 *
 * The dropped run is spelled `…`, which reads as a segment and matches a
 * disposition's `**` and nothing else.
 */
const withShortenedPath = (message: string): string | undefined => {
  const match = MODEL_PATH_RE.exec(message);
  if (match === null) {
    return undefined;
  }
  const [path] = match;
  const budget = MAX_MESSAGE_LENGTH - (message.length - path.length);
  const segments = path.split(".");
  const head = segments.at(0) ?? path;
  const tail = segments.slice(1);
  const spell = (keep: number): string =>
    [head, ELISION, ...tail.slice(tail.length - keep)].join(".");
  let kept = 0;
  while (kept < tail.length && spell(kept + 1).length <= budget) {
    kept += 1;
  }
  if (kept === 0) {
    return undefined;
  }
  return `${message.slice(0, match.index)}${spell(kept)}${message.slice(match.index + path.length)}`;
};

/**
 * Erase the parts of a message that vary per file: paths, identifiers, counts.
 *
 * Numbers go last and unconditionally. A message that distinguishes `numId 0`
 * from `numId 12` is reporting the same defect about two inputs, and the reserved
 * value that made `0` special is visible in the minimised reproduction, not here.
 */
export const normalizeFailureMessage = (raw: string): string => {
  const collapsed = raw.replaceAll(WHITESPACE_RE, " ").trim();
  const normalized = collapsed
    .replaceAll(GUID_RE, "<guid>")
    .replaceAll(ABSOLUTE_PATH_RE, "<path>")
    .replaceAll(HEX_RUN_RE, "<hex>")
    .replaceAll(DIGIT_RUN_RE, "N");
  if (normalized.length <= MAX_MESSAGE_LENGTH) {
    return normalized;
  }
  return withShortenedPath(normalized) ?? `${normalized.slice(0, MAX_MESSAGE_LENGTH)}${ELISION}`;
};

/**
 * The innermost stack frame inside a folio package, as `<package path>:<function>`.
 *
 * Frames in `node_modules` and in the gate's own scripts say nothing about which
 * folio code owns the defect, so they are skipped.
 */
export const topFolioFrame = (stack: string | undefined): string => {
  if (stack === undefined) {
    return NO_FRAME;
  }
  for (const line of stack.split("\n")) {
    const match = STACK_FRAME_RE.exec(line);
    if (match?.groups === undefined) {
      continue;
    }
    const { file, fn } = match.groups;
    if (file === undefined || file.includes("node_modules")) {
      continue;
    }
    const packageIndex = file.lastIndexOf("/packages/");
    if (packageIndex === -1) {
      continue;
    }
    const relative = file.slice(packageIndex + 1);
    return fn === undefined ? relative : `${relative}:${fn}`;
  }
  return NO_FRAME;
};

/** An error's message prefixed with its class, which is itself part of the defect. */
export const describeError = (cause: unknown): string => {
  if (!(cause instanceof Error)) {
    return String(cause);
  }
  const name = cause.name.length > 0 ? cause.name : cause.constructor.name;
  return cause.message.startsWith(`${name}:`) ? cause.message : `${name}: ${cause.message}`;
};

export const failureFromError = (invariant: CorpusInvariant, cause: unknown): CorpusFailure => ({
  invariant,
  message: normalizeFailureMessage(describeError(cause)),
  frame: topFolioFrame(cause instanceof Error ? cause.stack : undefined),
});

export const failureFromAssertion = (
  invariant: CorpusInvariant,
  message: string,
): CorpusFailure => ({
  invariant,
  message: normalizeFailureMessage(message),
  frame: NO_FRAME,
});

export const failureSignature = ({ invariant, message, frame }: CorpusFailure): string =>
  `${invariant} | ${message} | ${frame}`;

/**
 * One file's failures, at most one per signature.
 *
 * A census counts files per signature, so a file that exhibits the same
 * signature twice still counts once. Two failures collapsing to one signature
 * used to be rare; now that the model comparison reports every difference a
 * file exhibits rather than the first, it is ordinary — two paragraphs losing
 * the same field are one row, and a `files` count that reached 2 for one file
 * would ratchet against a defect nobody could shrink.
 */
export const distinctBySignature = (failures: readonly CorpusFailure[]): CorpusFailure[] => {
  const seen = new Set<string>();
  return failures.filter((failure) => {
    const signature = failureSignature(failure);
    if (seen.has(signature)) {
      return false;
    }
    seen.add(signature);
    return true;
  });
};
