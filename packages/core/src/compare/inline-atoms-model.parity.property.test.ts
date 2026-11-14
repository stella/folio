/**
 * The model matcher restores inline atoms exactly as the editor-state matcher
 * does.
 *
 * Both read the same story: a base whose text a comparison's own operations
 * have already rewritten toward a revised document, saved, so either side
 * opens the same bytes. Each runs through its reviewer's comparison access:
 * the editor-state matcher on a legacy reviewer, the model matcher on a
 * canonical one, whose batch commits on its session. The two results must
 * agree, and so must the two saved documents once every change is accepted,
 * and once every change is rejected: their text and the atoms each block
 * holds.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { relationshipIdOf } from "@stll/docx-core/model";

import { coverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FolioDocxReviewer,
  getFolioDocxComparisonAccess,
} from "../ai-edits/headless";
import { withComparisonNoteReferenceEdits } from "../ai-edits/note-references";
import {
  alignedNoteReferenceLabels,
  relabelFolioAIEditSnapshotNoteReferences,
  sourceDocumentOf,
} from "../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { fromMarkdown } from "../markdown";
import { createDocx, ensureParaIds } from "../server";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Footnote, ParagraphContent, RunContent } from "../types/document";
import { sameInlineAtoms } from "./inline-atoms";
import { planStoryCompare } from "./plan";

const MAIN = { type: "main" } as const;
const DATE = "2026-09-13T00:00:00.000Z";
/** The first revision id the comparison's own operations take. */
const ORIGINAL_SEED = 500;
const MAX_RANGES = 10_000;

const PICTURES = [
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
] as const;

const FIELD_TYPES = {
  " PAGE ": "PAGE",
  " NUMPAGES ": "NUMPAGES",
  " NUMWORDS ": "NUMWORDS",
} as const;
type FieldInstruction = keyof typeof FIELD_TYPES;
const FIELD_INSTRUCTIONS = Object.keys(FIELD_TYPES).filter(
  (instruction): instruction is FieldInstruction =>
    Object.hasOwn(FIELD_TYPES, instruction),
);

const FOOTNOTE_IDS = [10, 20] as const;

type Segment =
  | { kind: "text"; text: string }
  | {
      kind: "field";
      form: "simple" | "complex";
      instruction: FieldInstruction;
      result: string;
    }
  | { kind: "pageBreak" }
  | { kind: "picture"; picture: 0 | 1 }
  | { kind: "footnote"; id: (typeof FOOTNOTE_IDS)[number] };

type Body = readonly (readonly Segment[])[];

const run = (content: RunContent): ParagraphContent => ({
  type: "run",
  content: [content],
});

const contentOf = (segment: Segment): ParagraphContent => {
  switch (segment.kind) {
    case "text":
      return run({ type: "text", text: segment.text });
    case "field": {
      const result =
        segment.result === ""
          ? []
          : [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: segment.result }],
              },
            ];
      const fieldType = FIELD_TYPES[segment.instruction];
      return segment.form === "simple"
        ? {
            type: "simpleField",
            fieldType,
            instruction: segment.instruction,
            content: result,
          }
        : {
            type: "complexField",
            fieldType,
            instruction: segment.instruction,
            fieldCode: [],
            fieldResult: result,
          };
    }
    case "pageBreak":
      return run({ type: "break", breakType: "page" });
    case "picture":
      return run({
        type: "drawing",
        image: {
          type: "image",
          rId: relationshipIdOf(`rId_img_${String(segment.picture)}`),
          src: PICTURES[segment.picture],
          size: { width: 914_400, height: 914_400 },
          wrap: { type: "inline" },
        },
      });
    case "footnote":
      return run({ type: "footnoteRef", id: segment.id });
    default: {
      const unreachable: never = segment;
      return unreachable;
    }
  }
};

const buildDocx = async (body: Body): Promise<ArrayBuffer> => {
  const model = fromMarkdown(
    body.map((_, index) => `Paragraph ${String(index)}.`).join("\n\n"),
  );
  for (const [index, segments] of body.entries()) {
    const paragraph = model.package.document.content[index];
    if (paragraph?.type !== "paragraph")
      throw new Error("fixture paragraph is missing");
    paragraph.content = segments.map(contentOf);
  }
  model.package.footnotes = FOOTNOTE_IDS.map((id): Footnote => ({
    type: "footnote",
    id,
    noteType: "normal",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            content: [{ type: "text", text: `Footnote ${String(id)}.` }],
          },
        ],
      },
    ],
  }));
  const bytes = (await ensureParaIds(new Uint8Array(await createDocx(model))))
    .docx;
  return bytes.slice().buffer;
};

/** The bytes a canonical session reads, which the editor-state side reads too. */
const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const mainSnapshotOf = (
  reviewer: FolioDocxReviewer,
  mode: "with-revision-census" | "without-revision-census",
): FolioAIEditSnapshot => {
  const snapshot = getFolioDocxComparisonAccess(reviewer)
    .projectStories(mode)
    .stories.find(({ handle }) => handle.type === "main")?.snapshot;
  if (!snapshot) throw new Error("the main story has no snapshot");
  return snapshot;
};

type Prematch = {
  /** The base with the comparison's text operations applied, saved. */
  buffer: ArrayBuffer;
  nextRevisionId: number;
  labels: ReturnType<typeof alignedNoteReferenceLabels>;
};

/** The base as a comparison leaves it just before atoms are matched. */
const prematch = async (
  base: ArrayBuffer,
  target: ArrayBuffer,
): Promise<Prematch | null> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(base.slice(0));
  const baseSnapshot = mainSnapshotOf(reviewer, "with-revision-census");
  const targetSnapshot = mainSnapshotOf(
    await FolioDocxReviewer.fromBuffer(target.slice(0)),
    "without-revision-census",
  );
  const labels = alignedNoteReferenceLabels(
    sourceDocumentOf(baseSnapshot),
    sourceDocumentOf(targetSnapshot),
  );
  const plan = planStoryCompare({
    story: MAIN,
    baseSnapshot,
    targetSnapshot: relabelFolioAIEditSnapshotNoteReferences(
      targetSnapshot,
      labels,
    ),
    maxOperations: MAX_RANGES,
  });
  if (plan === null) return null;
  let nextRevisionId = ORIGINAL_SEED;
  if (plan.operations.length > 0) {
    const applied = withComparisonNoteReferenceEdits(() =>
      reviewer.applyDocumentOperationsToStory({
        story: MAIN,
        snapshot: baseSnapshot,
        revisionStamp: { date: DATE, idSeed: ORIGINAL_SEED },
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: plan.operations,
        },
        replacementBackground: "keep",
        undefinedReferences: "keep",
      }),
    );
    nextRevisionId = applied.nextRevisionId ?? ORIGINAL_SEED;
  }
  return { buffer: await reviewer.toBuffer(), nextRevisionId, labels };
};

const baseNotesOf = (reviewer: FolioDocxReviewer): ReadonlySet<string> =>
  new Set(
    reviewer
      .listStories()
      .flatMap(({ handle }) =>
        handle.type === "footnote" || handle.type === "endnote"
          ? [`${handle.type}:${String(handle.noteId)}`]
          : [],
      ),
  );

type MatchOutcome =
  | {
      status: "matched";
      nextRevisionId: number;
      changedTargetBlockIds: readonly string[];
      rangeCount: number;
      documentChanged: boolean;
    }
  | { status: "unalignable"; detail: string | undefined }
  | { status: "budget-exceeded" };

type SideResult = { outcome: MatchOutcome; buffer: ArrayBuffer };

type MatchResult = ReturnType<
  ReturnType<typeof getFolioDocxComparisonAccess>["matchInlineAtoms"]
>;

const outcomeOf = (result: MatchResult): MatchOutcome => {
  switch (result.status) {
    case "matched":
      return {
        status: "matched",
        nextRevisionId: result.nextRevisionId,
        changedTargetBlockIds: result.changedTargetBlockIds.toSorted(),
        rangeCount: result.rangeCount,
        documentChanged: result.documentChanged,
      };
    case "unalignable":
      return { status: "unalignable", detail: result.detail };
    case "budget-exceeded":
      return { status: "budget-exceeded" };
    default: {
      const unreachable: never = result;
      return unreachable;
    }
  }
};

type MatchOnOptions = {
  prematched: Prematch;
  target: ArrayBuffer;
  session: HeadlessReviewerSession;
};

/** The inline atom match on one engine, through the reviewer's comparison access. */
const matchOn = async ({
  prematched,
  target,
  session,
}: MatchOnOptions): Promise<SideResult> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    prematched.buffer.slice(0),
    { session },
  );
  const targetReviewer = await FolioDocxReviewer.fromBuffer(target.slice(0), {
    session,
  });
  const result = getFolioDocxComparisonAccess(reviewer).matchInlineAtoms({
    story: MAIN,
    targetReviewer,
    targetStory: MAIN,
    targetSnapshot: mainSnapshotOf(targetReviewer, "without-revision-census"),
    noteReferenceLabels: prematched.labels,
    baseNotes: baseNotesOf(reviewer),
    revisionStamp: { date: DATE, idSeed: prematched.nextRevisionId },
    originalRevisionIdSeed: ORIGINAL_SEED,
    maxRanges: MAX_RANGES,
  });
  return { outcome: outcomeOf(result), buffer: await reviewer.toBuffer() };
};

/** A saved document with every change accepted, then with every change rejected. */
const resolvedViews = async (buffer: ArrayBuffer) => {
  const accepting = await FolioDocxReviewer.fromBuffer(buffer.slice(0));
  accepting.acceptAll();
  const rejecting = await FolioDocxReviewer.fromBuffer(buffer.slice(0));
  rejecting.rejectAll();
  return [accepting.snapshot(), rejecting.snapshot()] as const;
};

type ParityCase = "unplanned" | "unchanged" | "restored" | "unalignable";

/** Run both matchers on one pair and hold them to the same result. */
const expectParity = async (
  baseBody: Body,
  targetBody: Body,
): Promise<ParityCase> => {
  const base = await ingested(await buildDocx(baseBody));
  const target = await ingested(await buildDocx(targetBody));
  const planned = await prematch(base, target);
  if (planned === null) return "unplanned";
  const prematched = { ...planned, buffer: await ingested(planned.buffer) };
  const legacy = await matchOn({ prematched, target, session: "legacy" });
  const canonical = await matchOn({
    prematched,
    target,
    session: "canonical",
  });
  expect(canonical.outcome).toEqual(legacy.outcome);
  const [legacyViews, canonicalViews] = [
    await resolvedViews(legacy.buffer),
    await resolvedViews(canonical.buffer),
  ];
  for (const [index, legacyView] of legacyViews.entries()) {
    const canonicalView = canonicalViews[index];
    if (!canonicalView) throw new Error("a resolved view is missing");
    expect(canonicalView.blocks.map(({ text }) => text)).toEqual(
      legacyView.blocks.map(({ text }) => text),
    );
    expect(
      sameInlineAtoms(
        sourceDocumentOf(canonicalView),
        sourceDocumentOf(legacyView),
      ),
    ).toBe(true);
  }
  switch (legacy.outcome.status) {
    case "matched":
      return legacy.outcome.documentChanged ? "restored" : "unchanged";
    case "unalignable":
    case "budget-exceeded":
      return "unalignable";
    default: {
      const unreachable: never = legacy.outcome;
      return unreachable;
    }
  }
};

const textArbitrary = fc.constantFrom(
  "Count: ",
  "Agreement applies to ",
  " today.",
  "12",
  "clause 1",
  " and ",
);

const segmentArbitrary: fc.Arbitrary<Segment> = fc.oneof(
  textArbitrary.map((text) => ({ kind: "text" as const, text })),
  fc.record({
    kind: fc.constant("field" as const),
    form: fc.constantFrom("simple" as const, "complex" as const),
    instruction: fc.constantFrom(...FIELD_INSTRUCTIONS),
    result: fc.constantFrom("1", "12", "clause 1", ""),
  }),
  fc.constant({ kind: "pageBreak" as const }),
  fc
    .constantFrom(0 as const, 1 as const)
    .map((picture) => ({ kind: "picture" as const, picture })),
  fc
    .constantFrom(...FOOTNOTE_IDS)
    .map((id) => ({ kind: "footnote" as const, id })),
);

/** How the revised document differs from the base, one step at a time. */
type Edit =
  | { kind: "drop"; paragraph: number; segment: number }
  | { kind: "insert"; paragraph: number; segment: number; value: Segment }
  | { kind: "fieldToText"; paragraph: number; segment: number }
  | { kind: "textToField"; paragraph: number; segment: number }
  | { kind: "merge"; paragraph: number }
  | { kind: "copy"; paragraph: number };

const editArbitrary: fc.Arbitrary<Edit> = fc.oneof(
  fc.record({
    kind: fc.constant("drop" as const),
    paragraph: fc.nat(),
    segment: fc.nat(),
  }),
  fc.record({
    kind: fc.constant("insert" as const),
    paragraph: fc.nat(),
    segment: fc.nat(),
    value: segmentArbitrary,
  }),
  fc.record({
    kind: fc.constant("fieldToText" as const),
    paragraph: fc.nat(),
    segment: fc.nat(),
  }),
  fc.record({
    kind: fc.constant("textToField" as const),
    paragraph: fc.nat(),
    segment: fc.nat(),
  }),
  fc.record({ kind: fc.constant("merge" as const), paragraph: fc.nat() }),
  fc.record({ kind: fc.constant("copy" as const), paragraph: fc.nat() }),
);

const applyEdit = (body: Body, edit: Edit): Body => {
  if (body.length === 0) return body;
  const at = edit.paragraph % body.length;
  const paragraph = body[at] ?? [];
  const replaced = (segments: readonly Segment[]) =>
    body.map((current, index) => (index === at ? segments : current));
  switch (edit.kind) {
    case "drop": {
      if (paragraph.length === 0) return body;
      const index = edit.segment % paragraph.length;
      return replaced(paragraph.filter((_, current) => current !== index));
    }
    case "insert": {
      const index = edit.segment % (paragraph.length + 1);
      return replaced([
        ...paragraph.slice(0, index),
        edit.value,
        ...paragraph.slice(index),
      ]);
    }
    case "fieldToText":
      return replaced(
        paragraph.map((segment, index) =>
          index === edit.segment % Math.max(paragraph.length, 1) &&
          segment.kind === "field"
            ? { kind: "text", text: segment.result }
            : segment,
        ),
      );
    case "textToField":
      return replaced(
        paragraph.map((segment, index) =>
          index === edit.segment % Math.max(paragraph.length, 1) &&
          segment.kind === "text"
            ? {
                kind: "field",
                form: "simple",
                instruction: " NUMWORDS ",
                result: segment.text,
              }
            : segment,
        ),
      );
    case "merge": {
      const next = body[at + 1];
      if (next === undefined) return body;
      return [
        ...body.slice(0, at),
        [...paragraph, ...next],
        ...body.slice(at + 2),
      ];
    }
    case "copy":
      return [...body.slice(0, at + 1), paragraph, ...body.slice(at + 1)];
    default: {
      const unreachable: never = edit;
      return unreachable;
    }
  }
};

const bodyArbitrary = fc.array(
  fc.array(segmentArbitrary, { minLength: 1, maxLength: 5 }),
  { minLength: 1, maxLength: 4 },
);

const pairArbitrary = fc
  .tuple(bodyArbitrary, fc.array(editArbitrary, { minLength: 1, maxLength: 4 }))
  .map(([base, edits]) => {
    let target: Body = base;
    for (const edit of edits) target = applyEdit(target, edit);
    return { base, target };
  });

const parityProperty = fc.asyncProperty(
  pairArbitrary,
  async ({ base, target }) => {
    await expectParity(base, target);
  },
);

const FIXED_PAIRS: readonly { name: string; base: Body; target: Body }[] = [
  {
    name: "a field replaces its materialized result",
    base: [[{ kind: "text", text: "Count: 12" }]],
    target: [
      [
        { kind: "text", text: "Count: " },
        {
          kind: "field",
          form: "simple",
          instruction: " NUMWORDS ",
          result: "12",
        },
      ],
    ],
  },
  {
    name: "retained result text replaces a field",
    base: [
      [
        { kind: "text", text: "Count: " },
        {
          kind: "field",
          form: "simple",
          instruction: " NUMWORDS ",
          result: "12",
        },
      ],
    ],
    target: [[{ kind: "text", text: "Count: 12" }]],
  },
  {
    name: "a page break and a picture come and go",
    base: [
      [
        { kind: "text", text: "Agreement applies to " },
        { kind: "pageBreak" },
        { kind: "text", text: " today." },
      ],
    ],
    target: [
      [
        { kind: "text", text: "Agreement applies to " },
        { kind: "picture", picture: 1 },
        { kind: "text", text: " today." },
      ],
    ],
  },
  {
    name: "a copied paragraph brings its footnote reference",
    base: [
      [
        { kind: "text", text: "Agreement applies to " },
        { kind: "footnote", id: 10 },
        { kind: "text", text: " today." },
      ],
      [{ kind: "text", text: "clause 1" }],
    ],
    target: [
      [
        { kind: "text", text: "Agreement applies to " },
        { kind: "footnote", id: 10 },
        { kind: "text", text: " today." },
      ],
      [{ kind: "text", text: "clause 1" }],
      [
        { kind: "text", text: "Agreement applies to " },
        { kind: "footnote", id: 10 },
        { kind: "text", text: " today." },
      ],
    ],
  },
  {
    name: "a field joins a merged paragraph",
    base: [
      [{ kind: "text", text: "Count: " }],
      [{ kind: "text", text: " today." }],
    ],
    target: [
      [
        { kind: "text", text: "Count: " },
        { kind: "field", form: "complex", instruction: " PAGE ", result: "1" },
        { kind: "text", text: " today." },
      ],
    ],
  },
];

describe("matchInlineAtomsOnModel", () => {
  for (const { name, base, target } of FIXED_PAIRS)
    test(
      `matches the editor-state matcher: ${name}`,
      async () => {
        await expectParity(base, target);
      },
      propertyTestTimeout(30_000),
    );

  test(
    "matches the editor-state matcher on generated document pairs",
    async () => {
      await assertProperty(parityProperty, {
        numRuns: 25,
        id: "model matcher parity with the editor-state matcher",
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "matches the editor-state matcher on a fixed sample that restores atoms",
    async () => {
      const cases: ParityCase[] = [];
      for (const { base, target } of coverageSample(pairArbitrary, 25))
        cases.push(await expectParity(base, target));
      expect(
        cases.filter((kind) => kind === "restored").length,
      ).toBeGreaterThanOrEqual(3);
    },
    propertyTestTimeout(120_000),
  );
});
