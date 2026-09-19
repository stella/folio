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

const MAX_MESSAGE_LENGTH = 160;

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
  return normalized.length > MAX_MESSAGE_LENGTH
    ? `${normalized.slice(0, MAX_MESSAGE_LENGTH)}…`
    : normalized;
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
