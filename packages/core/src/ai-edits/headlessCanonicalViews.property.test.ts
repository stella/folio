/**
 * A canonical headless reviewer's resolved views read as the legacy
 * reviewer's.
 *
 * For every story of a document, the original and final views (text,
 * blocks and the changes left) and the story's tables in each view are
 * read on both engines. A canonical reviewer reads them from a second
 * session resolved to the view, and leaves its own document as it was.
 * A story resolved to its final view keeps that view and holds no change.
 *
 * A note is part of a resolved view exactly when a reference to it survives
 * in that view, wherever the reference sits: in a run, a hyperlink, a content
 * control, a field or a table cell. A note whose every reference the view
 * removes leaves with it, as resolving the document removes it. A note no
 * reference reaches in the document itself stays in both views: resolving
 * removes only what a resolved change takes away.
 */

import { panic } from "better-result";
import { describe, expect, spyOn, test } from "bun:test";
import fc from "fast-check";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { isDeepStrictEqual } from "node:util";

import {
  applyDocumentOps,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  documentStories,
  leasedIdCensus,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { reviewerDocumentArbitrary } from "../../../../test/generators/reviewerDocumentArbitraries";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  rejectedContainerEndIdentities,
  rejectedJoinSurvivorIdentities,
  resolvedAbsorbedFollowerIdentities,
  withSurvivorProperties,
} from "../../../../test/support/rejectedJoinSurvivors";
import { rejectedParagraphFillerIdentities } from "../../../../test/support/rejectedParagraphFillers";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ProseMirrorDocumentValidationError } from "../prosemirror/validation";
import {
  CANONICAL_GAP,
  type HeadlessReviewerSession,
} from "../types/canonicalCapabilities";
import type {
  Document,
  Hyperlink,
  InlineSdt,
  Paragraph,
  ParagraphContent,
  Run,
  SimpleField,
} from "../types/document";
import {
  FOLIO_RESOLVED_REVIEWED_VIEWS,
  FolioDocxReviewer,
  type FolioResolvedReviewedView,
} from "./headless";

const FIXTURES_DIR = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus",
);

const FIXTURE_FILES: string[] = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
};

const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), { session });

/** Every story in both resolved views. */
const views = (reviewer: FolioDocxReviewer) =>
  reviewer.listStories().flatMap(({ handle }) =>
    FOLIO_RESOLVED_REVIEWED_VIEWS.map((view) => {
      const read = reviewer.readReviewedStory({ story: handle, view });
      return {
        handle,
        view,
        text: read?.text ?? null,
        changes: read?.changes ?? null,
        blocks: read?.snapshot.blocks ?? null,
        tables: reviewer
          .storyTables({ story: handle, view })
          .map((table) => table.node.textContent),
      };
    }),
  );

/** Every note a reference reaches anywhere in `value`, at any depth. */
const referencedNotes = (
  value: unknown,
  found: Set<string> = new Set(),
): Set<string> => {
  if (Array.isArray(value)) {
    for (const item of value) referencedNotes(item, found);
    return found;
  }
  if (typeof value !== "object" || value === null) return found;
  if (
    "type" in value &&
    "id" in value &&
    (value.type === "footnoteRef" || value.type === "endnoteRef") &&
    typeof value.id === "number"
  )
    found.add(
      `${value.type === "footnoteRef" ? "footnote" : "endnote"}:${String(value.id)}`,
    );
  for (const child of Object.values(value)) referencedNotes(child, found);
  return found;
};

/**
 * The legacy reviewer's views, less each note the view takes every reference
 * away from: the legacy reviewer reads a note's own state whether or not
 * anything still refers to it, while resolving the document removes it.
 */
const expectedViews = async (buffer: ArrayBuffer) => {
  const legacy = await open(buffer, "legacy");
  const referenced = referencedNotes(legacy.toDocument().package);
  const reached: Record<FolioResolvedReviewedView, Set<string>> = {
    original: referencedNotes(legacy.toResolvedDocument("reject").package),
    final: referencedNotes(legacy.toResolvedDocument("accept").package),
  };
  const expected = views(legacy);
  for (const read of expected)
    if (
      (read.handle.type === "footnote" || read.handle.type === "endnote") &&
      referenced.has(`${read.handle.type}:${String(read.handle.noteId)}`) &&
      !reached[read.view].has(
        `${read.handle.type}:${String(read.handle.noteId)}`,
      )
    )
      Object.assign(read, {
        text: null,
        changes: null,
        blocks: null,
        tables: [],
      });
  return expected;
};

const viewLaws = async (
  buffer: ArrayBuffer,
  unsupported: "fail" | "known-input-boundaries" = "fail",
): Promise<void> => {
  let canonical: FolioDocxReviewer;
  try {
    canonical = await open(buffer, "canonical");
  } catch (error) {
    if (
      unsupported !== "known-input-boundaries" ||
      !CanonicalSessionError.is(error)
    )
      throw error;
    // Generated wrappers can repeat bookmark endpoints. Require the shared
    // source converter to independently diagnose that malformed input; an
    // otherwise valid document's session refusal must fail this law.
    const parsed = await parseDocx(buffer, {
      detectVariables: false,
      preloadFonts: false,
    });
    let invalid: unknown;
    try {
      toProseDoc(parsed);
    } catch (cause) {
      invalid = cause;
    }
    if (!(invalid instanceof ProseMirrorDocumentValidationError)) throw error;
    expect(invalid.issues.length).toBeGreaterThan(0);
    for (const issue of invalid.issues)
      expect(issue.message).toMatch(
        /^Bookmark .+ has more than one (?:start|end) boundary \(paragraph .+\)\.$/u,
      );
    expect(error.message).toBe(
      `Canonical projection failed: ${invalid.message}`,
    );
    expect(error.reason).toBe("refused");
    expect(error.ground).toEqual({ gap: CANONICAL_GAP.dispatch });
    return;
  }
  const before = canonical.toDocument();
  let read: ReturnType<typeof views>;
  try {
    read = views(canonical);
  } catch (error) {
    // Generated input may place a reviewed paragraph before its container's
    // final table or a standing range boundary. Verify the exact native
    // refusal independently; an arbitrary session refusal must fail this law.
    // Fixed supported regressions use the default strict policy.
    if (
      unsupported !== "known-input-boundaries" ||
      !CanonicalSessionError.is(error)
    )
      throw error;
    const revisionIds = [...leasedIdCensus(before).revision];
    const stories = documentStories(before);
    const order = [
      ...stories.filter((story) => story !== OP_STORIES.MAIN),
      OP_STORIES.MAIN,
    ];
    const nativeRefusals = (["accept", "reject"] as const).flatMap(
      (decision) => {
        const resolved = applyDocumentOps(
          before,
          order.map((story) => ({
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story,
            revisionIds,
            decision,
          })),
        );
        if (resolved.isOk()) return [];
        const refusal = resolved.error;
        const knownBoundary =
          refusal.reason === DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK ||
          (refusal.reason === DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE &&
            (/^A range boundary stands between [0-9A-Fa-f]{8} and the next block\.$/u.test(
              refusal.message,
            ) ||
              /^Resolving the mark of [0-9A-Fa-f]{8} would leave its container without a final paragraph\.$/u.test(
                refusal.message,
              )));
        return knownBoundary &&
          refusal.opType === DOCUMENT_OP_TYPES.RESOLVE_REVISION
          ? [refusal.message]
          : [];
      },
    );
    expect(nativeRefusals).toContain(error.message);
    expect(error.reason).toBe("refused");
    expect(error.ground).toEqual({ gap: CANONICAL_GAP.dispatch });
    expect(canonical.toDocument()).toEqual(before);
    return;
  }
  const expected = await expectedViews(buffer);
  // The legacy reviewer keeps the wrong paragraph of a rejected join before a
  // block the joined paragraph cannot run into; the view must keep the later
  // one, whose mark stays.
  const survivors = rejectedJoinSurvivorIdentities({
    source: before,
    actual: canonical.toResolvedDocument("reject"),
    expected: (await open(buffer, "legacy")).toResolvedDocument("reject"),
  });
  // The join keeps the later paragraph's properties with its mark; the legacy
  // reviewer reads them on the earlier paragraph once it has them too.
  const aligned = withSurvivorProperties(before, survivors);
  if (aligned !== undefined) {
    const alignedBuffer = await packaged(aligned);
    if (alignedBuffer === null)
      panic("A document with a paragraph's properties copied packages.");
    const alignedViews = await expectedViews(alignedBuffer);
    for (const [index, entry] of expected.entries()) {
      if (entry.view !== "original") continue;
      expected[index] =
        alignedViews.find(
          (candidate) =>
            candidate.view === "original" &&
            isDeepStrictEqual(candidate.handle, entry.handle),
        ) ?? panic("The aligned document holds the same stories.");
    }
  }
  for (const entry of expected) {
    if (entry.view !== "original") continue;
    for (const { expectedId, survivorId } of survivors) {
      entry.blocks =
        entry.blocks?.map((block) =>
          block.id === expectedId ? { ...block, id: survivorId } : block,
        ) ?? null;
      entry.text =
        entry.text?.replaceAll(`[${expectedId}]`, `[${survivorId}]`) ?? null;
    }
  }
  // The legacy reviewer drops an empty paragraph whose rejected added mark
  // ends a local container; the paragraph stays there, as the canonical
  // reviewer reads it, after the block it follows. A control's last
  // paragraph whose mark either view resolves away takes an empty paragraph
  // after the control, which the legacy reviewer keeps.
  const legacyReviewer = await open(buffer, "legacy");
  const legacyRejected = legacyReviewer.toResolvedDocument("reject");
  const containerEnds = new Set(
    rejectedContainerEndIdentities({
      source: before,
      actual: canonical.toResolvedDocument("reject"),
      expected: legacyRejected,
    }),
  );
  const absorbedFollowers = {
    original: new Set(
      resolvedAbsorbedFollowerIdentities({
        source: before,
        actual: canonical.toResolvedDocument("reject"),
        expected: legacyRejected,
        decision: "reject",
      }),
    ),
    final: new Set(
      resolvedAbsorbedFollowerIdentities({
        source: before,
        actual: canonical.toResolvedDocument("accept"),
        expected: legacyReviewer.toResolvedDocument("accept"),
        decision: "accept",
      }),
    ),
  } satisfies Record<FolioResolvedReviewedView, Set<string>>;
  for (const entry of expected) {
    const absorbed = absorbedFollowers[entry.view];
    entry.blocks = entry.blocks?.filter(({ id }) => !absorbed.has(id)) ?? null;
    if (entry.view !== "original" || containerEnds.size === 0) continue;
    const actualBlocks =
      read.find(
        (candidate) =>
          candidate.view === "original" &&
          isDeepStrictEqual(candidate.handle, entry.handle),
      )?.blocks ?? [];
    for (const [index, block] of actualBlocks.entries()) {
      if (!containerEnds.has(block.id) || entry.blocks === null) continue;
      const previous = actualBlocks.at(index - 1);
      const at =
        index === 0 || previous === undefined
          ? 0
          : entry.blocks.findIndex(({ id }) => id === previous.id) + 1;
      if (at === 0 && index !== 0) continue;
      entry.blocks = entry.blocks.toSpliced(at, 0, block);
    }
  }
  const fillerIds = new Map(
    isDeepStrictEqual(read, expected)
      ? []
      : rejectedParagraphFillerIdentities({
          source: before,
          actual: canonical.toResolvedDocument("reject"),
          expected: (await open(buffer, "legacy")).toResolvedDocument("reject"),
        }).map(({ actualId, expectedId }) => [actualId, expectedId]),
  );
  expect(
    read.map((entry) =>
      entry.view !== "original"
        ? entry
        : Object.assign({}, entry, {
            blocks:
              entry.blocks?.map((block) =>
                Object.assign({}, block, {
                  id: fillerIds.get(block.id) ?? block.id,
                }),
              ) ?? null,
          }),
    ),
  ).toEqual(expected);
  // Reading a view resolves nothing in the reviewer itself.
  expect(canonical.toDocument()).toEqual(before);

  // Resolving one story reads, afterwards, as that story's view did.
  for (const { handle } of canonical.listStories()) {
    const reviewer = await open(buffer, "canonical");
    const final = reviewer.readReviewedStory({ story: handle, view: "final" });
    expect(
      reviewer.resolveReviewedStory({ story: handle, view: "final" }),
    ).toBe(true);
    // Its final view is unchanged, and no change of it is left pending.
    const resolved = reviewer.readReviewedStory({
      story: handle,
      view: "final",
    });
    expect(resolved?.text).toBe(final?.text);
    expect(
      reviewer.readReviewedStory({ story: handle, view: "current-markup" })
        ?.changes,
    ).toEqual([]);
  }
};

const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch {
    return null;
  }
};

// These malformed corpus documents fail session construction. Every other
// corpus document and every generated supported document must actually read.
const IMPORT_REFUSALS = new Map([
  [
    "sdt-self-closing.docx",
    {
      reason: "refused",
      ground: { gap: CANONICAL_GAP.storyContentProjection },
      message: "The projection changed paragraph structure.",
    },
  ],
]);

describe("a canonical reviewer's resolved views", () => {
  test.each(FIXTURE_FILES)(
    "%s reads its resolved views as the legacy reviewer does",
    async (filename) => {
      const buffer = await ingested(readFixture(filename));
      const refused = IMPORT_REFUSALS.get(filename);
      if (refused !== undefined) {
        await expect(open(buffer, "canonical")).rejects.toMatchObject(refused);
        return;
      }
      await viewLaws(buffer);
    },
    propertyTestTimeout(30_000),
  );

  test(
    "generated documents read their resolved views as the legacy reviewer does",
    async () => {
      await assertProperty(
        fc.asyncProperty(reviewerDocumentArbitrary, async (document) => {
          const buffer = await packaged(document);
          if (buffer !== null) await viewLaws(buffer, "known-input-boundaries");
        }),
        {
          id: "generated documents read their resolved views as the legacy reviewer does",
          numRuns: 25,
        },
      );
    },
    propertyTestTimeout(480_000),
  );
});

test("a typed refusal for a supported view fails the generated parity law", async () => {
  const buffer = await ingested(
    await createDocx(fromMarkdown("Supported source.")),
  );
  const read = spyOn(FolioDocxReviewer.prototype, "readReviewedStory");
  try {
    read.mockImplementation(() => {
      throw new CanonicalSessionError({
        reason: "refused",
        ground: { gap: CANONICAL_GAP.dispatch },
        message: "Unexpected supported-view refusal",
      });
    });
    await expect(viewLaws(buffer, "known-input-boundaries")).rejects.toThrow();
  } finally {
    read.mockRestore();
  }
});

for (const kind of ["ins", "del"] as const)
  for (const boundary of [
    { type: "bookmarkStart", id: 90, name: "_GoBack" },
    { type: "bookmarkEnd", id: 90 },
  ] as const)
    test(`a native ${boundary.type} boundary independently explains the ${kind} view refusal`, async () => {
      const buffer = await ingested(
        await createDocx({
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [],
                  pPrMark: { kind, info: { id: 1, author: "Reviewer" } },
                },
                boundary,
              ],
            },
          },
        }),
      );
      const reviewer = await open(buffer, "canonical");
      const before = reviewer.toDocument();
      const decision = kind === "ins" ? "reject" : "accept";
      const native = applyDocumentOps(before, [
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [1],
          decision,
        },
      ]);
      expect(native.isErr()).toBe(true);
      if (native.isOk())
        return panic("Expected the native range-boundary refusal");
      expect(native.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
      expect(native.error.opType).toBe(DOCUMENT_OP_TYPES.RESOLVE_REVISION);
      expect(native.error.message).toBe(
        "A range boundary stands between 00000001 and the next block.",
      );
      await expect(viewLaws(buffer)).rejects.toThrow(native.error.message);
      await viewLaws(buffer, "known-input-boundaries");
      expect(reviewer.toDocument()).toEqual(before);

      const read = spyOn(FolioDocxReviewer.prototype, "readReviewedStory");
      try {
        read.mockImplementation(() => {
          throw new CanonicalSessionError({
            reason: "refused",
            ground: { gap: CANONICAL_GAP.dispatch },
            message: "Unexpected refusal on a known-boundary input",
          });
        });
        await expect(
          viewLaws(buffer, "known-input-boundaries"),
        ).rejects.toThrow();
      } finally {
        read.mockRestore();
      }
    });

test.each(["ins", "del"] as const)(
  "a native final-table boundary independently explains the %s view refusal",
  async (kind) => {
    const buffer = await ingested(
      await createDocx({
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "00000001",
                content: [],
                pPrMark: { kind, info: { id: 1, author: "Reviewer" } },
              },
              {
                type: "table",
                rows: [
                  {
                    type: "tableRow",
                    cells: [
                      {
                        type: "tableCell",
                        content: [
                          {
                            type: "paragraph",
                            paraId: "00000002",
                            content: [],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        },
      }),
    );
    const reviewer = await open(buffer, "canonical");
    const before = reviewer.toDocument();
    const native = applyDocumentOps(before, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [1],
        decision: kind === "ins" ? "reject" : "accept",
      },
    ]);
    expect(native.isErr()).toBe(true);
    if (native.isOk())
      return panic("The final table cannot end the resolved container.");
    expect(native.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
    expect(native.error.opType).toBe(DOCUMENT_OP_TYPES.RESOLVE_REVISION);
    expect(native.error.message).toBe(
      "Resolving the mark of 00000001 would leave its container without a final paragraph.",
    );
    await expect(viewLaws(buffer)).rejects.toThrow(native.error.message);
    await viewLaws(buffer, "known-input-boundaries");
    expect(reviewer.toDocument()).toEqual(before);
  },
);

test.each(["ins", "del"] as const)(
  "a resolved %s paragraph before an atomic block preserves its section and source markup",
  async (kind) => {
    const document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [],
              sectionProperties: { pageWidth: 11906 },
            },
            {
              type: "paragraph",
              paraId: "00000002",
              content: [],
              pPrMark: { kind, info: { id: 1, author: "Reviewer" } },
            },
            { type: "preservedBlock", xml: "<w:altChunk/>" },
          ],
          finalSectionProperties: { pageWidth: 12240, pageHeight: 15840 },
        },
      },
    } satisfies Document;
    const buffer = await ingested(await createDocx(document));
    await viewLaws(buffer);
    for (const session of ["legacy", "canonical"] as const)
      for (const view of FOLIO_RESOLVED_REVIEWED_VIEWS) {
        const reviewer = await open(buffer, session);
        const before = reviewer.toDocument();
        const atomic = before.package.document.content.at(2);
        expect(atomic?.type).toBe("preservedBlock");
        const resolved = reviewer.toResolvedDocument(
          view === "final" ? "accept" : "reject",
        );
        const assertSourceStructure = (read: Document) => {
          const blocks = read.package.document.content;
          expect(blocks.map(({ type }) => type)).toEqual([
            "paragraph",
            "paragraph",
            "preservedBlock",
          ]);
          expect(blocks.at(0)).toMatchObject({
            type: "paragraph",
            paraId: "00000001",
            content: [],
            sectionProperties: { pageWidth: 11906 },
          });
          expect(blocks.at(1)).toMatchObject({
            type: "paragraph",
            paraId: "00000002",
            content: [],
          });
          const target = blocks.at(1);
          expect(
            target?.type === "paragraph" ? target.pPrMark : "missing",
          ).toBeUndefined();
          expect(blocks.at(2)).toEqual(atomic);
          expect(read.package.document.finalSectionProperties).toMatchObject({
            pageWidth: 12240,
            pageHeight: 15840,
          });
        };
        assertSourceStructure(resolved);
        expect(reviewer.toDocument()).toEqual(before);
        // Force semantic serialization, rather than replay the source ZIP.
        const saved = await createDocx({ package: resolved.package });
        assertSourceStructure((await open(saved, session)).toDocument());
      }
  },
  propertyTestTimeout(30_000),
);

const NOTE_ID = 37;
const NOTE = { type: "footnote", noteId: NOTE_ID } as const;
const REFERENCE: Run = {
  type: "run",
  content: [{ type: "footnoteRef", id: NOTE_ID }],
};

type Placement = "run" | "hyperlink" | "contentControl" | "field" | "tableCell";

/** Where the note's only reference sits. */
const PLACEMENTS: Record<Placement, Run | Hyperlink | InlineSdt | SimpleField> =
  {
    run: REFERENCE,
    hyperlink: {
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: [REFERENCE],
    } satisfies Hyperlink,
    contentControl: {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: 42, tag: "clause" },
      content: [REFERENCE],
    } satisfies InlineSdt,
    field: {
      type: "simpleField",
      instruction: "PAGE",
      fieldType: "PAGE",
      content: [REFERENCE],
    } satisfies SimpleField,
    tableCell: REFERENCE,
  };

/** How the reference is tracked, and the views it survives in. */
const TRACKING = {
  none: ["original", "final"],
  insertion: ["final"],
  deletion: ["original"],
} as const satisfies Record<string, readonly FolioResolvedReviewedView[]>;

const referenceParagraph = (
  placement: Placement,
  tracking: keyof typeof TRACKING,
): Paragraph => {
  const carrier = PLACEMENTS[placement];
  const info = { id: 91, author: "A" };
  if (tracking === "none") return paragraph([run("See "), carrier]);
  const content: ParagraphContent =
    tracking === "insertion"
      ? { type: "insertion", info, content: [carrier] }
      : { type: "deletion", info, content: [carrier] };
  return paragraph([run("See "), content]);
};

const placedDocument = (
  placement: Placement,
  tracking: keyof typeof TRACKING,
): Document => {
  const document = fromMarkdown("Body.");
  const host = referenceParagraph(placement, tracking);
  document.package.document.content.push(
    placement === "tableCell"
      ? {
          type: "table",
          columnWidths: [4800],
          rows: [
            {
              type: "tableRow",
              cells: [{ type: "tableCell", content: [host] }],
            },
          ],
        }
      : host,
    paragraph("After."),
  );
  document.package.footnotes = [
    { type: "footnote", id: NOTE_ID, content: [paragraph("Note text.")] },
  ];
  return document;
};

describe("a note in a canonical reviewer's resolved views", () => {
  for (const placement of Object.keys(PLACEMENTS) as Placement[]) {
    for (const tracking of Object.keys(TRACKING) as Array<
      keyof typeof TRACKING
    >) {
      test(
        `a reference in a ${placement}, ${tracking}: the note reads in each view its reference survives`,
        async () => {
          const buffer = await ingested(
            await createDocx(placedDocument(placement, tracking)),
          );
          const reviewer = await open(buffer, "canonical");
          const survives: readonly FolioResolvedReviewedView[] =
            TRACKING[tracking];
          for (const view of FOLIO_RESOLVED_REVIEWED_VIEWS) {
            const text =
              reviewer.readReviewedStory({ story: NOTE, view })?.text ?? null;
            if (survives.includes(view)) expect(text).toContain("Note text.");
            else expect(text).toBeNull();
          }
          // The reviewer's own document keeps the note whatever the views do.
          expect(
            reviewer.readReviewedStory({ story: NOTE, view: "current-markup" })
              ?.text,
          ).toContain("Note text.");
        },
        propertyTestTimeout(30_000),
      );
    }
  }
});

describe("text a resolved paragraph mark joins into another paragraph", () => {
  test.each(["ins", "del", "moveFrom", "moveTo"] as const)(
    "%s keeps run and field formatting separate from the surviving mark",
    async (kind) => {
      for (const wrapper of ["run", "field"] as const)
        for (const styled of [false, true])
          for (const direct of [false, true]) {
            const document = fromMarkdown("x");
            document.package.styles = {
              docDefaults: { rPr: { fontSize: 22 } },
              styles: [
                { type: "paragraph", styleId: "Normal", default: true },
                {
                  type: "paragraph",
                  styleId: "Source",
                  basedOn: "Normal",
                  rPr: { fontSize: 20 },
                },
                {
                  type: "paragraph",
                  styleId: "Target",
                  basedOn: "Normal",
                  rPr: { fontSize: 28 },
                },
              ],
            };
            const mark = (id: number) => ({
              kind,
              info: { id, author: "Reviewer" },
            });
            const text = {
              ...run("a"),
              ...(direct ? { formatting: { fontSize: 30 } } : {}),
            };
            const payload =
              wrapper === "run"
                ? text
                : ({
                    type: "complexField",
                    instruction: "DATE",
                    fieldType: "DATE",
                    fieldCode: [
                      {
                        type: "run",
                        content: [{ type: "instrText", text: " DATE " }],
                      },
                    ],
                    fieldResult: [text],
                  } satisfies ParagraphContent);
            const targetStyle = styled ? "Target" : "Normal";
            document.package.document.content = [
              { type: "paragraph", content: [], pPrMark: mark(1) },
              {
                type: "paragraph",
                content: [payload],
                formatting: { styleId: "Source" },
                pPrMark: mark(2),
              },
              // The empty mark's caret glyph is8pt; incoming body runs are not.
              {
                type: "paragraph",
                content: [],
                formatting: {
                  styleId: targetStyle,
                  runProperties: { fontSize: 16 },
                },
              },
              {
                type: "paragraph",
                content: [run("b")],
                formatting: { styleId: targetStyle },
              },
            ];
            const buffer = await ingested(await createDocx(document));
            const view =
              kind === "ins" || kind === "moveTo" ? "original" : "final";
            for (const session of ["legacy", "canonical"] as const) {
              const reviewer = await open(buffer, session);
              const before = reviewer.toDocument();
              const read = reviewer.readReviewedStory({
                story: { type: "main" },
                view,
              });
              const sizeOf = (wantedText: string) => {
                const block = read?.snapshot.blocks.find(
                  (candidate) =>
                    "previewRuns" in candidate &&
                    candidate.previewRuns?.at(0)?.text === wantedText,
                );
                return block !== undefined && "previewRuns" in block
                  ? block.previewRuns?.at(0)?.fontSizePt
                  : undefined;
              };
              expect(sizeOf("b")).toBe(styled ? 14 : 11);
              expect(sizeOf("a")).toBe(direct ? 15 : sizeOf("b"));
              expect(reviewer.toDocument()).toEqual(before);
              const resolved = reviewer.toResolvedDocument(
                view === "final" ? "accept" : "reject",
              );
              const surviving = resolved.package.document.content.at(0);
              expect(
                surviving?.type === "paragraph"
                  ? surviving.formatting?.runProperties?.fontSize
                  : "missing",
              ).toBe(16);
              const reopened = await open(
                await createDocx({ package: resolved.package }),
                session,
              );
              expect(
                reopened.readReviewedStory({ story: { type: "main" }, view })
                  ?.snapshot.blocks,
              ).toEqual(read?.snapshot.blocks);
            }
          }
    },
    propertyTestTimeout(30_000),
  );
});

// Seed 2113650910: once it passes, its path replays the whole shrink tree,
// which outlasts the property's budget, so the shrunk case is fixed here.
test("a comment range across a content control holding only a linked bookmark reads its resolved views", async () => {
  const commentBody = (id: number, paraId: string) => ({
    id,
    author: "A",
    content: [{ type: "paragraph" as const, paraId, content: [] }],
  });
  const buffer = await ingested(
    await createDocx({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [{ type: "commentRangeStart", id: 34 }],
            },
            {
              type: "blockSdt",
              properties: { sdtType: "group", id: 1 },
              content: [
                {
                  type: "paragraph",
                  paraId: "00000002",
                  content: [
                    {
                      type: "inlineSdt",
                      properties: { sdtType: "richText", id: 2, tag: "clause" },
                      content: [
                        {
                          type: "hyperlink",
                          rId: "rId9",
                          href: "https://example.org",
                          children: [
                            { type: "bookmarkStart", id: 1, name: "_Link" },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
            {
              type: "paragraph",
              paraId: "00000003",
              content: [{ type: "commentRangeEnd", id: 34 }],
              sectionProperties: { pageWidth: 11906 },
            },
          ],
          finalSectionProperties: { pageWidth: 12240, pageHeight: 15840 },
          comments: [commentBody(7, "7FFFFFF0"), commentBody(34, "00000004")],
        },
      },
    }),
  );
  await viewLaws(buffer);
});
