import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { hashFolioAIBlockText } from "../ai-edits/snapshot";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const PROPERTY_BUDGET = 120_000;
const BLOCK_ID = "12345678";

setDefaultTimeout(propertyTestTimeout(PROPERTY_BUDGET));

// Words that share leading and trailing letters, so the smallest change of
// one into another keeps some of them ("defined" to "updated" keeps "ed").
const VOCABULARY = [
  "remarks",
  "revised",
  "defined",
  "updated",
  "thanks",
  "done",
  "order",
  "signed",
  "invoiced",
  "period",
] as const;

const SHAPES = {
  separate: "separate",
  commentFirst: "comment-first",
  replaceFirst: "replace-first",
} as const;
type Shape = (typeof SHAPES)[keyof typeof SHAPES];
type Mode = "direct" | "tracked-changes";

type TextSpan = { start: number; end: number };

type Case = {
  words: readonly string[];
  /** What the comment covers and what the edit replaces, in the joined text. */
  comment: TextSpan;
  replaced: TextSpan;
  replacement: string;
};

/** Where each word of `words`, joined by single spaces, starts and ends. */
const wordSpans = (words: readonly string[]): TextSpan[] => {
  const spans: TextSpan[] = [];
  let offset = 0;
  for (const word of words) {
    spans.push({ start: offset, end: offset + word.length });
    offset += word.length + 1;
  }
  return spans;
};

/** Whole words, or any characters: a comment can start inside a word. */
const textSpan = (words: readonly string[]): fc.Arbitrary<TextSpan> => {
  const spans = wordSpans(words);
  const length = words.join(" ").length;
  const ordered = ([left, right]: [number, number]) => ({
    first: Math.min(left, right),
    last: Math.max(left, right),
  });
  return fc.oneof(
    fc
      .tuple(
        fc.nat({ max: words.length - 1 }),
        fc.nat({ max: words.length - 1 }),
      )
      .map(ordered)
      .map(({ first, last }) => ({
        start: spans[first]?.start ?? panic("A generated word must exist"),
        end: spans[last]?.end ?? panic("A generated word must exist"),
      })),
    fc
      .tuple(fc.nat({ max: length - 1 }), fc.nat({ max: length - 1 }))
      .map(ordered)
      .map(({ first, last }) => ({ start: first, end: last + 1 })),
  );
};

const cases: fc.Arbitrary<Case> = fc
  .array(fc.constantFrom(...VOCABULARY), { minLength: 2, maxLength: 4 })
  .chain((words) =>
    fc.record({
      words: fc.constant(words),
      comment: textSpan(words),
      replaced: textSpan(words),
      replacement: fc
        .array(fc.constantFrom(...VOCABULARY), { minLength: 1, maxLength: 2 })
        .map((replacement) => replacement.join(" ")),
    }),
  )
  .filter(
    ({ words, replaced, replacement }) =>
      words.join(" ").slice(replaced.start, replaced.end) !== replacement,
  );

const sourceBuffer = async (text: string): Promise<ArrayBuffer> => {
  const document = fromMarkdown(text);
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph")
    return panic("The generated paragraph is missing");
  paragraph.paraId = BLOCK_ID;
  const prepared = await prepareCanonicalDocxInput(await createDocx(document));
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    session: "canonical",
    author: "Reviewer",
  });

const commentsOf = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getComments()
    .map(({ text, anchoredText }) => ({ text, anchoredText }));

type Outcome = {
  /** The comments as the edit leaves them, tracked changes pending. */
  live: ReturnType<typeof commentsOf>;
  /** The comments once every tracked change is accepted. */
  resolved: ReturnType<typeof commentsOf>;
  /** Those of the saved live document, reopened. */
  liveReopened: ReturnType<typeof commentsOf>;
  /** Those of the saved resolved document, reopened. */
  resolvedReopened: ReturnType<typeof commentsOf>;
};

type RunOptions = {
  buffer: ArrayBuffer;
  mode: Mode;
  shape: Shape;
  input: Case;
};

const run = async ({
  buffer,
  mode,
  shape,
  input,
}: RunOptions): Promise<Outcome | "overlapping"> => {
  const text = input.words.join(" ");
  const { comment: covered, replaced } = input;
  const comment: FolioDocumentOperation = {
    id: "comment",
    type: "commentOnRange",
    range: {
      type: "textRange",
      story: "main",
      blockId: BLOCK_ID,
      startOffset: covered.start,
      endOffset: covered.end,
      selectedTextHash: hashFolioAIBlockText(
        text.slice(covered.start, covered.end),
      ),
    },
    comment: { text: "Note" },
  };
  const replace: FolioDocumentOperation = {
    id: "replace",
    type: "replaceRange",
    range: {
      type: "textRange",
      story: "main",
      blockId: BLOCK_ID,
      startOffset: replaced.start,
      endOffset: replaced.end,
      selectedTextHash: hashFolioAIBlockText(
        text.slice(replaced.start, replaced.end),
      ),
    },
    replace: input.replacement,
  };
  const batches = (() => {
    switch (shape) {
      case SHAPES.separate:
        return [[comment], [replace]];
      case SHAPES.commentFirst:
        return [[comment, replace]];
      case SHAPES.replaceFirst:
        return [[replace, comment]];
      default: {
        const unreachable: never = shape;
        return panic(`Unknown batch shape ${String(unreachable)}`);
      }
    }
  })();
  const reviewer = await open(buffer);
  for (const operations of batches) {
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        atomic: true,
        operations,
      },
      { undefinedReferences: "refuse" },
    );
    // A batch admits a comment beside its edit, not one cutting across it.
    if (
      shape !== SHAPES.separate &&
      result.skipped.some(({ reason }) => reason === "overlappingOperation")
    )
      return "overlapping";
    expect(result.skipped).toEqual([]);
  }
  const live = commentsOf(reviewer);
  const liveReopened = commentsOf(await open(await reviewer.toBuffer()));
  if (mode === "tracked-changes") reviewer.acceptAll();
  const resolved = commentsOf(reviewer);
  const resolvedReopened = commentsOf(await open(await reviewer.toBuffer()));
  return { live, resolved, liveReopened, resolvedReopened };
};

test("a comment goes with its words whatever the batch and mode", async () => {
  await assertProperty(
    fc.asyncProperty(cases, async (input) => {
      const text = input.words.join(" ");
      const buffer = await sourceBuffer(text);
      const outcomes = new Map<string, Outcome>();
      for (const mode of ["direct", "tracked-changes"] as const)
        for (const shape of Object.values(SHAPES)) {
          const outcome = await run({ buffer, mode, shape, input });
          if (outcome !== "overlapping")
            outcomes.set(`${mode} ${shape}`, outcome);
        }
      const reference =
        outcomes.get(`direct ${SHAPES.separate}`) ??
        panic("The reference outcome must run");
      for (const [label, outcome] of outcomes) {
        // A batch's comment stands before its edit runs, so the batch equals
        // the comment and then the edit; accepting a tracked edit equals it
        // made directly.
        expect({ label, resolved: outcome.resolved }).toEqual({
          label,
          resolved: reference.resolved,
        });
        // Every live comment saves and reopens as the reviewer shows it.
        expect({ label, comments: outcome.liveReopened }).toEqual({
          label,
          comments: outcome.live,
        });
        expect({ label, comments: outcome.resolvedReopened }).toEqual({
          label,
          comments: outcome.resolved,
        });
      }
      const tracked = [...outcomes]
        .filter(([label]) => label.startsWith("tracked-changes"))
        .map(([, { live }]) => live);
      for (const live of tracked) expect(live).toEqual(tracked[0]);
      // A comment every character of which the edit takes, with no letter
      // left to keep, retires; one with words outside the edit keeps them.
      const { comment: covered, replaced } = input;
      const coveredText = text.slice(covered.start, covered.end);
      const keptBefore = text.slice(
        covered.start,
        Math.max(covered.start, Math.min(covered.end, replaced.start)),
      );
      const keptAfter = text.slice(
        Math.min(covered.end, Math.max(covered.start, replaced.end)),
        covered.end,
      );
      const takenWhole =
        replaced.start <= covered.start && covered.end <= replaced.end;
      if (
        takenWhole &&
        ![...coveredText].some((character) =>
          input.replacement.includes(character),
        )
      )
        expect(reference.resolved).toEqual([]);
      if (!takenWhole) {
        const anchored = (reference.resolved.at(0)?.anchoredText ?? "").trim();
        expect(reference.resolved).toHaveLength(1);
        expect(anchored.startsWith(keptBefore.trim())).toBe(true);
        expect(anchored.endsWith(keptAfter.trim())).toBe(true);
      }
    }),
    {
      id: "a comment goes with its words whatever the batch and mode",
      numRuns: 20,
    },
  );
});
