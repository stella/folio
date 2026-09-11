import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import type { Mark, Node as PMNode } from "prosemirror-model";
import { history } from "prosemirror-history";
import {
  EditorState,
  Plugin,
  PluginKey,
  TextSelection,
  type Command,
  type Transaction,
} from "prosemirror-state";
import { ySyncPlugin, yUndoPlugin } from "y-prosemirror";
import * as Y from "yjs";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import type { FolioAIBlock } from "../../ai-edits/types";
import { parseDocx } from "../../docx/parser";
import { createDocx, createEmptyDocx, repackDocx } from "../../docx/rezip";
import {
  pluginsForHeadlessRevisionResolution,
  stateAllowsHeadlessRevisionResolution,
} from "../../internal/headlessRevisionResolutionGuard";
import type {
  Document,
  HeaderFooter,
  Paragraph,
  StyleDefinitions,
  TextFormatting,
  TrackedRunChange,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import {
  getChangedParagraphIds,
  hasStructuralChanges,
  hasUntrackedChanges,
  ParagraphChangeTrackerExtension,
} from "../extensions/features/ParagraphChangeTrackerExtension";
import { createDocumentStylesPlugin } from "../plugins/documentStyles";
import { schema } from "../schema";
import {
  acceptChange,
  acceptAllChanges,
  rejectChange,
  rejectAllChanges,
  resolveAllChangesInHeadlessState,
} from "./comments";

const AUTHOR = "Reviewer";
const DATE = "2026-09-09T00:00:00.000Z";
const HEADER_RELATIONSHIP_ID = "rIdBulkResolutionHeader";

const persistedBlockProjection = (block: FolioAIBlock): Omit<FolioAIBlock, "idStability"> => {
  const persisted = { ...block };
  delete persisted.idStability;
  return persisted;
};

const persistedSnapshotProjection = (blocks: readonly FolioAIBlock[] | undefined) =>
  blocks?.map(persistedBlockProjection);

const trackerExtension = ParagraphChangeTrackerExtension();
const changeTrackerPlugin = trackerExtension.onSchemaReady({ schema }).plugins?.at(0);
if (!changeTrackerPlugin) {
  throw new Error("Expected the paragraph change tracker plugin.");
}

const stepCountKey = new PluginKey<number>("headlessBulkResolutionStepCount");
const stepCountPlugin = new Plugin<number>({
  key: stepCountKey,
  state: {
    init: () => 0,
    apply: (transaction, previous) =>
      transaction.docChanged ? transaction.steps.length : previous,
  },
});

const serializedStepsKey = new PluginKey<readonly unknown[]>("headlessBulkSerializedSteps");
const serializedStepsPlugin = new Plugin<readonly unknown[]>({
  key: serializedStepsKey,
  state: {
    init: () => [],
    apply: (transaction, previous) =>
      transaction.docChanged ? transaction.steps.map((step) => step.toJSON()) : previous,
  },
});

type AppliedCommand = {
  state: EditorState;
  transaction: Transaction;
};

const apply = (state: EditorState, command: Command): AppliedCommand => {
  let transaction: Transaction | null = null;
  expect(
    command(state, (dispatched) => {
      transaction = dispatched;
    }),
  ).toBe(true);
  expect(transaction).not.toBeNull();
  if (!transaction) {
    throw new Error("Expected the revision command to dispatch.");
  }
  return { state: state.apply(transaction), transaction };
};

const revisionMark = (
  type: "insertion" | "deletion",
  revisionId: number,
  moveKind?: "moveTo" | "moveFrom",
): Mark =>
  schema.marks[type].create({
    revisionId,
    author: AUTHOR,
    date: DATE,
    ...(moveKind ? { moveKind } : {}),
  });

type MarkedTextRunsOptions = {
  type: "insertion" | "deletion";
  idOffset: number;
};

const markedTextRuns = ({ type, idOffset }: MarkedTextRunsOptions): PMNode[] =>
  Array.from({ length: 300 }, (_, index) =>
    schema.text(`tracked-${index} `, [revisionMark(type, idOffset + index)]),
  );

const paragraphMark = (kind: "ins" | "del" | "moveTo" | "moveFrom", id: number) => ({
  kind,
  info: { id, author: AUTHOR, date: DATE },
});

const paragraph = (
  seed: number,
  index: number,
  nextId: () => number,
  nestedInline: "include" | "omit",
): PMNode => {
  const insertionId = nextId();
  const deletionId = nextId();
  const runPropertyId = nextId();
  const comment = schema.marks.comment.create({ commentId: seed * 100 + index });
  const runProperty = schema.marks.runPropertyChange.create({
    changes: [
      {
        type: "runPropertyChange",
        info: { id: runPropertyId, author: AUTHOR, date: DATE },
        previousFormatting: { italic: true },
        currentFormatting: { bold: true },
      },
    ],
  });
  const fieldHyperlink = schema.marks.hyperlink.create({
    href: `#field-${seed}`,
    anchor: `field-${seed}`,
    _docxHyperlinkIndex: seed,
  });
  const content = [
    schema.text(`old-${seed}-${index}`, [
      revisionMark("deletion", deletionId, index % 5 === 0 ? "moveFrom" : undefined),
    ]),
    ...(index === 0
      ? [schema.nodes.bookmarkBoundary.create({ type: "start", id: seed, name: `seed-${seed}` })]
      : []),
    schema.text(` anchor-${index} `, [comment]),
    ...(index === 0 ? [schema.nodes.bookmarkBoundary.create({ type: "end", id: seed })] : []),
    schema.text(`new-${seed}-${index}`, [
      revisionMark("insertion", insertionId, index % 5 === 0 ? "moveTo" : undefined),
    ]),
    schema.text(` format-${index}`, [schema.marks.bold.create(), runProperty]),
    schema.nodes.hardBreak.create({ breakType: index % 2 === 0 ? "column" : null }, null, [
      revisionMark(index % 2 === 0 ? "insertion" : "deletion", nextId()),
    ]),
    schema.text(` alternate-old-${index}`, [revisionMark("deletion", nextId())]),
    schema.text(` alternate-new-${index}`, [revisionMark("insertion", nextId())]),
    schema.text(` alternate-format-${index}`, [runProperty]),
    ...(index === 0 && nestedInline === "include"
      ? [
          schema.node(
            "sdt",
            {
              tag: `nested-${seed}`,
              rawPropertiesXml: `<w:sdtPr><w:tag w:val="nested-${seed}"/></w:sdtPr>`,
            },
            [
              schema.text("sdt-old", [revisionMark("deletion", nextId())]),
              schema.node(
                "structuredField",
                {
                  fieldType: "REF",
                  instruction: `REF nested_${seed}`,
                  displayText: "",
                },
                [
                  schema.text("field-old", [fieldHyperlink, revisionMark("deletion", nextId())]),
                  schema.nodes.bookmarkBoundary.create(
                    {
                      type: "start",
                      id: seed + 1000,
                      name: `field-${seed}`,
                    },
                    null,
                    [fieldHyperlink],
                  ),
                  schema.text("field-new", [fieldHyperlink, revisionMark("insertion", nextId())]),
                  schema.nodes.bookmarkBoundary.create({ type: "end", id: seed + 1000 }, null, [
                    fieldHyperlink,
                  ]),
                ],
              ),
              schema.text("sdt-new", [revisionMark("insertion", nextId())]),
            ],
          ),
        ]
      : []),
  ];
  const attrs: Record<string, unknown> = {
    paraId: `${(seed * 100 + index).toString(16).padStart(8, "0")}`,
  };
  if (index % 3 === 0) {
    attrs["alignment"] = "right";
    attrs["_originalFormatting"] = { alignment: "right" };
    attrs["_propertyChanges"] = [
      {
        type: "paragraphPropertyChange",
        info: { id: nextId(), author: AUTHOR, date: DATE },
        previousFormatting: { alignment: "left", keepNext: true },
      },
    ];
  }
  if (index % 4 === 1) {
    attrs["pPrMark"] = paragraphMark(index % 2 === 0 ? "ins" : "del", nextId());
  }
  if (index === 2) {
    attrs["sectionBreakType"] = "continuous";
    attrs["_sectionProperties"] = {
      sectionStart: "continuous",
      pageSize: { width: 12_240, height: 15_840 },
      propertyChanges: [
        {
          type: "sectionPropertyChange",
          info: { id: nextId(), author: AUTHOR, date: DATE },
          previousProperties: {
            sectionStart: "nextPage",
            pageSize: { width: 11_900, height: 16_800 },
          },
        },
      ],
    };
  }
  return schema.node("paragraph", attrs, content);
};

const nestedTable = (seed: number, nextId: () => number): PMNode => {
  const emptyRunProperty = schema.marks.runPropertyChange.create({ changes: [] });
  const nested = schema.node("table", null, [
    schema.node("tableRow", null, [
      schema.node("tableCell", null, [
        schema.node("paragraph", null, [
          schema.text("nested-old", [revisionMark("deletion", nextId())]),
          schema.text("nested-new", [revisionMark("insertion", nextId())]),
        ]),
      ]),
    ]),
  ]);
  const cellPropertyId = nextId();
  const cellMarkerId = nextId();
  const firstCell = schema.node(
    "tableCell",
    {
      backgroundColor: "99CCFF",
      tcPrChange: [
        {
          info: { id: cellPropertyId, author: AUTHOR, date: DATE },
          previousFormatting: { backgroundColor: "FFFFFF", verticalAlign: "top" },
        },
      ],
    },
    [schema.node("paragraph"), nested, schema.node("paragraph", null, schema.text("tail"))],
  );
  const secondCell = schema.node(
    "tableCell",
    {
      cellMarker: {
        kind: seed % 2 === 0 ? "ins" : "del",
        info: { revisionId: cellMarkerId, author: AUTHOR, date: DATE },
      },
    },
    [schema.node("paragraph", null, schema.text("membership"))],
  );
  const changedRow = schema.node(
    "tableRow",
    {
      height: 480,
      trPrChange: [
        {
          info: { id: nextId(), author: AUTHOR, date: DATE },
          previousFormatting: { height: 240, isHeader: true },
        },
      ],
      ...(seed % 3 === 0 ? { trIns: { revisionId: nextId(), author: AUTHOR, date: DATE } } : {}),
    },
    [firstCell, secondCell],
  );
  const stableRow = schema.node("tableRow", null, [
    schema.node("tableCell", null, [
      schema.node("paragraph", null, schema.text("stable row", [emptyRunProperty])),
    ]),
  ]);
  return schema.node(
    "table",
    {
      width: 7200,
      tblPrChange: [
        {
          info: { id: nextId(), author: AUTHOR, date: DATE },
          previousFormatting: { width: 6400, justification: "center" },
        },
      ],
    },
    [stableRow, changedRow],
  );
};

const generatedDocument = (
  seed: number,
  paragraphCount = 8,
  nestedInline: "include" | "omit" = "include",
): PMNode => {
  let revisionId = seed * 1000 + 1;
  const nextId = () => revisionId++;
  const blocks: PMNode[] = [];
  for (let index = 0; index < paragraphCount; index++) {
    blocks.push(paragraph(seed, index, nextId, nestedInline));
    if (index === 3) {
      blocks.push(nestedTable(seed, nextId));
    }
  }
  return schema.node("doc", null, blocks);
};

const BULK_RUN_PROPERTY_STYLES = {
  styles: [
    {
      type: "paragraph",
      styleId: "BulkParagraph",
      rPr: { bold: true, fontFamily: { ascii: "Aptos", hAnsi: "Aptos" } },
    },
    {
      type: "character",
      styleId: "BulkCharacter",
      rPr: { italic: true, fontSize: 28 },
    },
  ],
} as const satisfies StyleDefinitions;

type BulkRunPropertyFixtureKind = "styled" | "unstyled";

type BulkRunPropertyFixture = {
  document: Document;
  expected: {
    accept: TextFormatting;
    reject: TextFormatting;
  };
};

const bulkRunPropertyFixture = (
  carrierCount: number,
  kind: BulkRunPropertyFixtureKind,
): BulkRunPropertyFixture => {
  const document = createEmptyDocument();
  const styled = kind === "styled";
  const previousFormatting: TextFormatting = styled
    ? { styleId: "BulkCharacter", color: { rgb: "008000" } }
    : { bold: true };
  const currentFormatting: TextFormatting = styled
    ? { styleId: "BulkCharacter", color: { rgb: "800000" } }
    : { italic: true };
  const paragraphs = Array.from(
    { length: carrierCount },
    (_, index): Paragraph => ({
      type: "paragraph",
      ...(styled ? { formatting: { styleId: "BulkParagraph" } } : {}),
      content: [
        {
          type: "run",
          formatting: currentFormatting,
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: { id: index + 1, author: AUTHOR, date: DATE },
              previousFormatting,
              currentFormatting,
            },
          ],
          content: [{ type: "text", text: `carrier-${index}` }],
        },
      ],
    }),
  );
  document.package.styles = styled ? BULK_RUN_PROPERTY_STYLES : { styles: [] };
  document.package.document.content = paragraphs;
  return {
    document,
    expected: { accept: currentFormatting, reject: previousFormatting },
  };
};

const modelRunFormatting = (document: Document): (TextFormatting | undefined)[] =>
  document.package.document.content.map((block) => {
    if (block.type !== "paragraph") {
      throw new Error("Expected a paragraph-only bulk run-property fixture.");
    }
    const run = block.content.at(0);
    if (run?.type !== "run") {
      throw new Error("Expected one run in each bulk run-property paragraph.");
    }
    return run.formatting;
  });

const modelRunPropertyChangeCount = (document: Document): number => {
  let count = 0;
  for (const block of document.package.document.content) {
    if (block.type !== "paragraph") {
      continue;
    }
    for (const child of block.content) {
      if (child.type === "run") {
        count += child.propertyChanges?.length ?? 0;
      }
    }
  }
  return count;
};

const pmRunPropertyChangeCount = (doc: PMNode): number => {
  let count = 0;
  doc.descendants((node) => {
    count += node.marks.filter(({ type }) => type.name === "runPropertyChange").length;
  });
  return count;
};

const secondaryStoryBuffer = async (): Promise<ArrayBuffer> => {
  const document = await parseDocx(await createEmptyDocx(), {
    detectVariables: false,
    preloadFonts: false,
  });
  const content: TrackedRunChange[] = Array.from({ length: 300 }, (_, index) => ({
    type: "insertion",
    info: { id: index + 1, author: AUTHOR, date: DATE },
    content: [{ type: "run", content: [{ type: "text", text: `item-${index} ` }] }],
  }));
  const header = {
    type: "header",
    hdrFtrType: "default",
    content: [{ type: "paragraph", paraId: "70000001", content }],
  } satisfies HeaderFooter;
  document.package.headers = new Map([[HEADER_RELATIONSHIP_ID, header]]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER_RELATIONSHIP_ID }],
  };
  return repackDocx(document, { updateModifiedDate: false });
};

describe("headless bulk revision resolution equivalence", () => {
  test.each([
    { mode: "reject", revisionType: "insertion" },
    { mode: "accept", revisionType: "deletion" },
  ] as const)(
    "$mode fits an emptied required inline parent with the legacy position map",
    ({ mode, revisionType }) => {
      const field = schema.node(
        "structuredField",
        {
          fieldType: "REF",
          instruction: "REF required_inline_parent",
          displayText: "removed",
        },
        [schema.text("removed", [revisionMark(revisionType, 1)])],
      );
      const trailingText = "following";
      const doc = schema.node("doc", null, [
        schema.node("paragraph", null, [field, schema.text(trailingText)]),
      ]);
      const trailingTextStart = 1 + field.nodeSize;
      const selection = TextSelection.create(doc, trailingTextStart + 2);
      const state = EditorState.create({
        schema,
        doc,
        selection,
        plugins: pluginsForHeadlessRevisionResolution([stepCountPlugin]),
      });
      const bulk = resolveAllChangesInHeadlessState(state, mode);
      const legacy = apply(
        state,
        mode === "accept" ? acceptChange(0, doc.content.size) : rejectChange(0, doc.content.size),
      ).state;

      const legacyField = legacy.doc.firstChild?.firstChild;
      expect(field.childCount).toBe(1);
      expect(legacy.doc.textContent).toBe(trailingText);
      expect(legacyField?.type.name).toBe("structuredField");
      expect(legacyField?.firstChild?.type.name).toBe("tab");
      expect(bulk.doc.toJSON()).toEqual(legacy.doc.toJSON());
      expect(stepCountKey.getState(bulk)).toBe(1);
      expect({ anchor: bulk.selection.anchor, head: bulk.selection.head }).toEqual({
        anchor: legacy.selection.anchor,
        head: legacy.selection.head,
      });
      expect(() => legacy.doc.check()).not.toThrow();
      expect(() => bulk.doc.check()).not.toThrow();
    },
  );

  test("matches the legacy small-document semantics across generated nested revisions", () => {
    fc.assert(
      fc.property(fc.integer({ min: 1, max: 100_000 }), (seed) => {
        const doc = generatedDocument(seed, 4);

        for (const mode of ["accept", "reject"] as const) {
          const state = EditorState.create({
            schema,
            doc,
            plugins: pluginsForHeadlessRevisionResolution([changeTrackerPlugin]),
          });
          const bulkState = resolveAllChangesInHeadlessState(state, mode);
          const legacy = apply(
            state,
            mode === "accept"
              ? acceptChange(0, doc.content.size)
              : rejectChange(0, doc.content.size),
          );

          expect(bulkState.doc.toJSON()).toEqual(legacy.state.doc.toJSON());
          const bulkTracker = {
            paraIds: [...getChangedParagraphIds(bulkState)].toSorted(),
            structural: hasStructuralChanges(bulkState),
            untracked: hasUntrackedChanges(bulkState),
          };
          const legacyTracker = {
            paraIds: [...getChangedParagraphIds(legacy.state)].toSorted(),
            structural: hasStructuralChanges(legacy.state),
            untracked: hasUntrackedChanges(legacy.state),
          };
          expect(bulkTracker).toEqual(legacyTracker);
        }
      }),
      { seed: 2_609_090, numRuns: 24, verbose: true },
    );
  });

  for (const kind of ["unstyled", "styled"] as const) {
    for (const mode of ["accept", "reject"] as const) {
      test.each([255, 256, 257])(
        `${mode} matches the legacy run-property reconstruction for ${kind} inputs at %i carriers`,
        async (carrierCount) => {
          const fixture = bulkRunPropertyFixture(carrierCount, kind);
          const doc = toProseDoc(fixture.document);
          expect(pmRunPropertyChangeCount(doc)).toBe(carrierCount);
          const state = EditorState.create({
            schema,
            doc,
            plugins: [
              ...pluginsForHeadlessRevisionResolution([changeTrackerPlugin, stepCountPlugin]),
              createDocumentStylesPlugin(fixture.document.package.styles),
            ],
          });

          const bulk = resolveAllChangesInHeadlessState(state, mode);
          const legacy = apply(
            state,
            mode === "accept"
              ? acceptChange(0, doc.content.size)
              : rejectChange(0, doc.content.size),
          ).state;

          expect(bulk.doc.toJSON()).toEqual(legacy.doc.toJSON());
          expect(pmRunPropertyChangeCount(bulk.doc)).toBe(0);
          expect(stepCountKey.getState(bulk)).toBe(1);

          const bulkModel = fromProseDoc(bulk.doc, fixture.document);
          const legacyModel = fromProseDoc(legacy.doc, fixture.document);
          expect(bulkModel.package.document.content).toEqual(legacyModel.package.document.content);
          expect(modelRunPropertyChangeCount(bulkModel)).toBe(0);
          expect(modelRunFormatting(bulkModel)).toEqual(
            Array.from({ length: carrierCount }, () => fixture.expected[mode]),
          );

          const reopened = await parseDocx(await createDocx(bulkModel), {
            detectVariables: false,
            preloadFonts: false,
          });
          expect(modelRunPropertyChangeCount(reopened)).toBe(0);
          expect(modelRunFormatting(reopened)).toEqual(
            Array.from({ length: carrierCount }, () => fixture.expected[mode]),
          );
        },
      );
    }
  }

  test("keeps the bulk transaction structurally bounded as the document grows", () => {
    const stepCounts = [4, 400, 4_000].map((paragraphCount) => {
      const insertion = revisionMark("insertion", 1);
      const deletion = revisionMark("deletion", 2);
      const doc = schema.node(
        "doc",
        null,
        Array.from({ length: paragraphCount }, (_, index) =>
          schema.node("paragraph", null, [
            schema.text(`old-${index}`, [deletion]),
            schema.text(`new-${index}`, [insertion]),
          ]),
        ),
      );
      const state = EditorState.create({
        schema,
        doc,
        plugins: pluginsForHeadlessRevisionResolution([stepCountPlugin]),
      });
      const accepted = resolveAllChangesInHeadlessState(state, "accept");
      const rejected = resolveAllChangesInHeadlessState(state, "reject");
      return {
        accept: stepCountKey.getState(accepted),
        reject: stepCountKey.getState(rejected),
      };
    });

    expect(stepCounts).toEqual([
      { accept: 1, reject: 1 },
      { accept: 1, reject: 1 },
      { accept: 1, reject: 1 },
    ]);

    const doc = schema.node(
      "doc",
      null,
      Array.from({ length: 400 }, (_, index) =>
        schema.node("paragraph", null, [
          schema.text(`old-${index}`, [revisionMark("deletion", 1)]),
          schema.text(`new-${index}`, [revisionMark("insertion", 2)]),
        ]),
      ),
    );
    const editorState = EditorState.create({ schema, doc });
    expect(apply(editorState, acceptAllChanges()).transaction.steps).toHaveLength(800);
    expect(apply(editorState, rejectAllChanges()).transaction.steps).toHaveLength(800);
  });
});

describe("headless bulk revision resolution isolation", () => {
  test.each(["accept", "reject"] as const)(
    "%s falls back to serializable legacy steps when history is present",
    (mode) => {
      const paragraphCount = 300;
      const insertion = revisionMark("insertion", 1);
      const deletion = revisionMark("deletion", 2);
      const doc = schema.node(
        "doc",
        null,
        Array.from({ length: paragraphCount }, (_, index) =>
          schema.node("paragraph", null, [
            schema.text(`old-${index}`, [deletion]),
            schema.text(`new-${index}`, [insertion]),
          ]),
        ),
      );
      const state = EditorState.create({
        schema,
        doc,
        plugins: [
          ...pluginsForHeadlessRevisionResolution([stepCountPlugin, serializedStepsPlugin]),
          history(),
        ],
      });

      expect(stateAllowsHeadlessRevisionResolution(state)).toBe(false);
      const resolved = resolveAllChangesInHeadlessState(state, mode);
      const legacy = apply(
        state,
        mode === "accept" ? acceptChange(0, doc.content.size) : rejectChange(0, doc.content.size),
      );

      expect(resolved.doc.toJSON()).toEqual(legacy.state.doc.toJSON());
      expect(stepCountKey.getState(resolved)).toBe(paragraphCount * 2);
      expect(serializedStepsKey.getState(resolved)).toHaveLength(paragraphCount * 2);
    },
  );

  test("rejects actual synchronization and undo plugins from the headless state boundary", () => {
    const yDoc = new Y.Doc();
    const state = EditorState.create({
      schema,
      doc: generatedDocument(73),
      plugins: [
        ...pluginsForHeadlessRevisionResolution([]),
        ySyncPlugin(yDoc.getXmlFragment("prosemirror")),
        yUndoPlugin(),
      ],
    });

    expect(stateAllowsHeadlessRevisionResolution(state)).toBe(false);
  });

  test("rejects a collaboration-keyed plugin from the headless state boundary", () => {
    const state = EditorState.create({
      schema,
      doc: generatedDocument(74),
      plugins: [
        ...pluginsForHeadlessRevisionResolution([]),
        new Plugin({ key: new PluginKey("collab") }),
      ],
    });

    expect(stateAllowsHeadlessRevisionResolution(state)).toBe(false);
  });
});

describe("headless bulk revision state", () => {
  test.each(["accept", "reject"] as const)(
    "%s reports mark-only paragraphs to selective save without a structural fallback",
    (mode) => {
      const revision = revisionMark(mode === "accept" ? "insertion" : "deletion", 9);
      const doc = schema.node("doc", null, [
        schema.node("paragraph", { paraId: "changed" }, [
          schema.text("tracked ", [revision]),
          ...markedTextRuns({
            type: mode === "accept" ? "insertion" : "deletion",
            idOffset: 10_000,
          }),
        ]),
        schema.node("paragraph", { paraId: "stable" }, schema.text("stable")),
      ]);
      const state = EditorState.create({
        schema,
        doc,
        plugins: pluginsForHeadlessRevisionResolution([changeTrackerPlugin]),
      });
      const resolved = resolveAllChangesInHeadlessState(state, mode);

      expect([...getChangedParagraphIds(resolved)]).toEqual(["changed"]);
      expect(hasStructuralChanges(resolved)).toBe(false);
    },
  );
});

describe("bulk revision lifecycle", () => {
  test.each(["accept", "reject"] as const)(
    "%s resolves a bulk secondary story and persists it",
    async (mode) => {
      const story = {
        type: "header",
        relationshipId: HEADER_RELATIONSHIP_ID,
      } as const;
      const reviewer = await FolioDocxReviewer.fromBuffer(await secondaryStoryBuffer());
      expect(reviewer.readReviewedStory({ story, view: "current-markup" })?.changes.length).toBe(
        300,
      );

      if (mode === "accept") {
        reviewer.acceptAll();
      } else {
        reviewer.rejectAll();
      }
      const expected = reviewer.readReviewedStory({ story, view: "current-markup" });
      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      const actual = reopened.readReviewedStory({ story, view: "current-markup" });

      expect(actual?.changes).toEqual([]);
      expect(actual?.text).toBe(expected?.text);
    },
  );

  test.each(["accept", "reject"] as const)(
    "%s survives save and reopen with no unresolved revision carriers",
    async (mode) => {
      const doc = generatedDocument(mode === "accept" ? 71 : 72, 8, "omit");
      const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(fromProseDoc(doc)));
      const before = reviewer.readReviewedStory({ view: "current-markup" });
      expect(before?.changes.length).toBeGreaterThan(0);

      if (mode === "accept") {
        reviewer.acceptAll();
      } else {
        reviewer.rejectAll();
      }
      const expected = reviewer.readReviewedStory({ view: "current-markup" });
      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      const actual = reopened.readReviewedStory({ view: "current-markup" });

      expect(actual?.changes).toEqual([]);
      expect(
        expected?.snapshot.blocks.some(({ idStability }) => idStability === "positional"),
      ).toBe(true);
      expect(actual?.snapshot.blocks.every(({ idStability }) => idStability === undefined)).toBe(
        true,
      );
      expect(persistedSnapshotProjection(actual?.snapshot.blocks)).toEqual(
        persistedSnapshotProjection(expected?.snapshot.blocks),
      );
      expect(actual?.text).toBe(expected?.text);
    },
  );

  test.each(["accept", "reject"] as const)(
    "%s preserves nested inline containers and boundary atoms through save and reopen",
    async (mode) => {
      const seed = mode === "accept" ? 71 : 72;
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await createDocx(fromProseDoc(generatedDocument(seed))),
      );
      if (mode === "accept") {
        reviewer.acceptAll();
      } else {
        reviewer.rejectAll();
      }
      const expected = reviewer.readReviewedStory({ view: "current-markup" });
      const saved = await reviewer.toBuffer();
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      const actual = reopened.readReviewedStory({ view: "current-markup" });
      const roundTripped = toProseDoc(
        await parseDocx(saved, { detectVariables: false, preloadFonts: false }),
      );
      let inlineSdt: PMNode | null = null;
      let structuredField: PMNode | null = null;
      const bookmarkBoundaries: PMNode[] = [];
      roundTripped.descendants((node) => {
        if (node.type.name === "sdt") {
          inlineSdt = node;
        } else if (node.type.name === "structuredField") {
          structuredField = node;
        } else if (node.type.name === "bookmarkBoundary") {
          bookmarkBoundaries.push(node);
        }
      });

      expect(actual?.changes).toEqual([]);
      expect(actual?.text).toBe(expected?.text);
      expect(inlineSdt?.attrs["tag"]).toBe(`nested-${seed}`);
      expect(inlineSdt?.attrs["rawPropertiesXml"]).toContain(`<w:tag w:val="nested-${seed}"/>`);
      expect(structuredField?.attrs["instruction"]).toContain(`REF nested_${seed}`);
      expect(inlineSdt?.textContent).toContain(mode === "accept" ? "sdt-new" : "sdt-old");
      expect(
        bookmarkBoundaries
          .filter((node) => node.attrs["id"] === seed + 1000)
          .map((node) => node.attrs["type"]),
      ).toEqual(["start", "end"]);
    },
  );

  test.each(["accept", "reject"] as const)(
    "%s persists a mark-only resolution through selective save and reopen",
    async (mode) => {
      const revision = revisionMark(mode === "accept" ? "insertion" : "deletion", 91);
      const doc = schema.node("doc", null, [
        schema.node("paragraph", { paraId: "00000091" }, [
          schema.text("tracked ", [revision]),
          ...markedTextRuns({
            type: mode === "accept" ? "insertion" : "deletion",
            idOffset: 20_000,
          }),
        ]),
        schema.node("paragraph", { paraId: "00000092" }, schema.text("stable")),
      ]);
      const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(fromProseDoc(doc)));

      if (mode === "accept") {
        reviewer.acceptAll();
      } else {
        reviewer.rejectAll();
      }
      const expected = reviewer.readReviewedStory({ view: "current-markup" });
      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      const story = reopened.readReviewedStory({ view: "current-markup" });

      expect(story?.changes).toEqual([]);
      expect(story?.text).toBe(expected?.text);
    },
  );
});
