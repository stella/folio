import { expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import {
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  documentStories,
  OP_STORIES,
  paragraphLogicalText,
  REVISION_DECISIONS,
  storyBody,
  type DocumentOp,
  type EditorIntent,
  type OpStory,
} from "@stll/docx-core/ops";

import { PARAGRAPH_MARK_CHANGE_KINDS } from "@stll/docx-core/model";
import { parseDocumentBody } from "./documentParser";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";

import type { BlockContent, Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";
import { serializeDocument } from "./serializer/documentSerializer";
import { buildStructuralDocumentPatch } from "./structuralXmlPatch";

const apply = (document: Document, intent: EditorIntent) => {
  const ids = createEditorIntentIdAllocator()(document, intent);
  const planned = compileEditorIntent(document, {
    intent,
    mode: {
      type: "suggesting",
      revision: { id: ids.revisionId, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
      newIds: ids.newIds,
    },
  });
  if (planned.isErr()) throw planned.error;
  const applied = applyDocumentOps(document, planned.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const applyWithInverse = (document: Document, intent: EditorIntent) => {
  const ids = createEditorIntentIdAllocator()(document, intent);
  const planned = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: ids.newIds },
  });
  if (planned.isErr()) throw planned.error;
  const applied = applyDocumentOps(document, planned.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value;
};

// Exercise edit sequences and both save paths; a static serializer fixture misses
// metadata synthesized only when paragraph marks and structural splices are added.
test.each(["absent", "empty", "bold"] as const)(
  "pending edits retain exact content through full and structural saves (%s mark properties)",
  async (markProperties) => {
    const seed = createEmptyDocument({ initialText: "A😀B" });
    seed.package.document.content.push(
      ...createEmptyDocument({ initialText: "C" }).package.document.content,
    );
    for (const [index, paragraph] of seed.package.document.content.entries()) {
      if (paragraph.type !== "paragraph") panic("Expected paragraphs");
      paragraph.paraId = (index + 17).toString(16).padStart(8, "0").toUpperCase();
    }
    if (markProperties !== "absent") {
      for (const paragraph of seed.package.document.content) {
        if (paragraph.type !== "paragraph") panic("Expected paragraphs");
        paragraph.formatting = {
          ...paragraph.formatting,
          runProperties: markProperties === "empty" ? {} : { bold: true },
        };
      }
    }
    const source = await createDocx(seed);
    const original = await parseDocx(source, { preloadFonts: false, detectVariables: false });
    const first = original.package.document.content.at(0);
    const second = original.package.document.content.at(1);
    if (
      first?.type !== "paragraph" ||
      second?.type !== "paragraph" ||
      !first.paraId ||
      !second.paraId
    )
      panic("Expected identified paragraphs");
    const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });
    const inserted = apply(original, {
      type: "replaceText",
      from: at(first.paraId, 4),
      to: at(first.paraId, 4),
      text: "x",
    });
    const deleted = apply(inserted, {
      type: "replaceText",
      from: at(first.paraId, 1),
      to: at(first.paraId, 3),
      text: "",
    });
    const splitIds = createEditorIntentIdAllocator()(deleted, {
      type: "splitParagraph",
      at: at(first.paraId, 1),
    });
    const split = apply(deleted, {
      type: "splitParagraph",
      at: at(first.paraId, 1),
      newBlockId: splitIds.newBlockId,
    });
    const suggested = apply(split, {
      type: "joinParagraphs",
      story: OP_STORIES.MAIN,
      blockId: first.paraId,
      nextBlockId: second.paraId,
    });
    const fullSave = await createDocx(suggested);
    const zip = await JSZip.loadAsync(source);
    const originalXml = await zip.file("word/document.xml")?.async("text");
    if (!originalXml) panic("Expected document part");
    const patched = buildStructuralDocumentPatch({
      originalXml,
      serializedXml: serializeDocument(suggested),
      changedIds: new Set([first.paraId, second.paraId, splitIds.newBlockId]),
    });
    if (patched === null) panic("Expected structural save to handle the edit sequence");
    zip.file("word/document.xml", patched);
    const structuralSave = await zip.generateAsync({ type: "arraybuffer" });
    for (const saved of [fullSave, structuralSave]) {
      const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
      expect(reopened.package.document.content).toStrictEqual(suggested.package.document.content);
      for (const decision of Object.values(REVISION_DECISIONS)) {
        const revisionIds = getTrackedChangeStatsFromDoc(toProseDoc(suggested)).ids;
        expect(revisionIds.length).toBeGreaterThan(0);
        const resolution = {
          type: "resolveRevision",
          story: OP_STORIES.MAIN,
          decision,
          revisionIds,
        } as const;
        const before = applyDocumentOps(suggested, [resolution]);
        const after = applyDocumentOps(reopened, [resolution]);
        if (before.isErr()) throw before.error;
        if (after.isErr()) throw after.error;
        expect(after.value.document.package.document.content).toStrictEqual(
          before.value.document.package.document.content,
        );
        expect(getTrackedChangeStatsFromDoc(toProseDoc(after.value.document)).ids).toEqual([]);
      }
    }
  },
);

// Static story fixtures never exercised the two closed-slice insertion seams.
// Vary authored property presence, run boundaries and insertion positions in every story.
test.each(
  ["absent", "empty", "bold"].flatMap((runProperties) =>
    ["absent", "present"].map((runBoundary) => [runProperties, runBoundary] as const),
  ),
)(
  "text seam edits across every story are save/reopen fixed points and invertible (%s formatting, %s run boundary)",
  async (runProperties, runBoundary) => {
    const createStoryParagraph = ({ paraId, text }: { paraId: string; text: string }) => ({
      type: "paragraph" as const,
      paraId,
      content: [
        {
          type: "run" as const,
          ...(runBoundary === "present"
            ? { preservedAttributes: [{ name: "rsidR", value: "00A1B2C3" }] }
            : {}),
          ...(runProperties === "absent"
            ? {}
            : { formatting: runProperties === "empty" ? {} : { bold: true } }),
          content: [{ type: "text" as const, text }],
        },
      ],
    });
    const seed = createEmptyDocument({ initialText: "main" });
    seed.package.document.content = [createStoryParagraph({ paraId: "00000011", text: "main" })];
    seed.package.headers = new Map([
      [
        "rIdCanonicalHeader",
        {
          type: "header",
          hdrFtrType: "default",
          content: [createStoryParagraph({ paraId: "00000012", text: "header" })],
        },
      ],
    ]);
    seed.package.footers = new Map([
      [
        "rIdCanonicalFooter",
        {
          type: "footer",
          hdrFtrType: "default",
          content: [createStoryParagraph({ paraId: "00000013", text: "footer" })],
        },
      ],
    ]);
    seed.package.footnotes = [
      {
        type: "footnote",
        id: 2,
        content: [createStoryParagraph({ paraId: "00000014", text: "footnote" })],
      },
    ];
    seed.package.endnotes = [
      {
        type: "endnote",
        id: 3,
        content: [createStoryParagraph({ paraId: "00000015", text: "endnote" })],
      },
    ];
    seed.package.document.finalSectionProperties = {
      ...seed.package.document.finalSectionProperties,
      headerReferences: [{ type: "default", rId: "rIdCanonicalHeader" }],
      footerReferences: [{ type: "default", rId: "rIdCanonicalFooter" }],
    };

    const source = await createDocx(seed);
    const original = await parseDocx(source, { preloadFonts: false, detectVariables: false });
    const stories = documentStories(original).map((story) => {
      const paragraph = storyParagraph(original, story);
      return {
        story,
        id: paragraph.paraId ?? panic("Expected an identified story paragraph"),
        text: paragraphLogicalText(paragraph),
      };
    });
    for (const position of ["start", "interior", "end"] as const) {
      let edited = original;
      const inverses: DocumentOp[][] = [];
      for (const { story, id, text } of stories) {
        const offset = { start: 0, interior: 2, end: text.length }[position];
        const applied = applyWithInverse(edited, {
          type: "replaceText",
          from: { story, blockId: id, offset },
          to: { story, blockId: id, offset },
          text: "X",
        });
        edited = applied.document;
        inverses.push(applied.inverse);
      }
      const reopened = await parseDocx(await createDocx(edited), {
        preloadFonts: false,
        detectVariables: false,
      });
      for (const { story, id } of stories) {
        const paragraph = storyParagraph(reopened, story);
        expect(paragraph.paraId).toBe(id);
        expect(paragraph).toStrictEqual(storyParagraph(edited, story));
      }
      const restored = applyDocumentOps(reopened, inverses.toReversed().flat());
      if (restored.isErr()) throw restored.error;
      for (const { story } of stories) {
        const paragraph = storyParagraph(restored.value.document, story);
        expect(paragraph.content).toStrictEqual(storyParagraph(original, story).content);
      }
    }
  },
);

const storyParagraph = (document: Document, story: OpStory) => {
  const paragraph = storyBody(document, story).content.find((block) => block.type === "paragraph");
  if (!paragraph || paragraph.type !== "paragraph") panic("Expected original story paragraph");
  return paragraph;
};

test("every paragraph mark kind retains absent, empty and populated formatting", () => {
  for (const kind of [undefined, ...PARAGRAPH_MARK_CHANGE_KINDS]) {
    for (const properties of [undefined, {}, { bold: true }]) {
      for (const runInWithNext of [undefined, false, true]) {
        const original = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000011",
                  formatting: {
                    styleId: "Normal",
                    ...(runInWithNext === undefined ? {} : { runInWithNext }),
                    ...(properties === undefined ? {} : { runProperties: properties }),
                  },
                  ...(kind === undefined
                    ? {}
                    : { pPrMark: { kind, info: { id: 1, author: "Reviewer" } } }),
                  content: [
                    { type: "run", formatting: {}, content: [{ type: "text", text: "é😀" }] },
                  ],
                },
              ],
            },
          },
        } satisfies Document;
        const serialized = serializeDocument(original);
        for (const xml of [
          serialized,
          serialized
            .replaceAll("folio:emptyMarkProperties", "f:emptyMarkProperties")
            .replaceAll("xmlns:folio=", "xmlns:f="),
        ]) {
          const reopened = parseDocumentBody(xml);
          expect(reopened.content).toStrictEqual(original.package.document.content);
          expect(
            parseDocumentBody(serializeDocument({ package: { document: reopened } })).content,
          ).toStrictEqual(reopened.content);
        }
      }
    }
  }
});

test.each(["body", "cell"] as const)(
  "repeated paragraph formatting and joins fold to one save-stable review in a %s",
  async (container) => {
    const paragraphs = [
      {
        type: "paragraph",
        paraId: "00000011",
        formatting: { alignment: "center", styleId: "Heading1" },
        content: [{ type: "run", content: [{ type: "text", text: "alpha" }] }],
      },
      {
        type: "paragraph",
        paraId: "00000012",
        formatting: { alignment: "end", styleId: "Normal" },
        propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: {
              id: 8,
              author: "Original author",
              date: "2026-01-01T00:00:00Z",
              initials: "OA",
            },
            previousFormatting: { alignment: "start", styleId: "Normal" },
          },
        ],
        content: [{ type: "run", content: [{ type: "text", text: "omega" }] }],
      },
    ] satisfies Paragraph[];
    const content =
      container === "body"
        ? paragraphs
        : ([
            {
              type: "table",
              rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: paragraphs }] }],
            },
          ] satisfies BlockContent[]);
    const seed = createEmptyDocument();
    seed.package.document.content = content;
    const original = await parseDocx(await createDocx(seed), {
      preloadFonts: false,
      detectVariables: false,
    });
    const repeated = applyDocumentOps(original, [
      {
        type: "setParagraphProps",
        story: OP_STORIES.MAIN,
        blockId: "00000012",
        patch: { alignment: "right" },
        revision: {
          id: 100,
          author: "First reviewer",
          date: "2026-02-01T00:00:00Z",
          initials: "FR",
        },
      },
      {
        type: "setParagraphProps",
        story: OP_STORIES.MAIN,
        blockId: "00000012",
        patch: { styleId: "Heading2" },
        revision: { id: 101, author: "Second reviewer", date: "2026-02-02T00:00:00Z" },
      },
    ]).unwrap();
    const intent = {
      type: "joinParagraphs",
      story: OP_STORIES.MAIN,
      blockId: "00000011",
      nextBlockId: "00000012",
    } as const;
    const ids = createEditorIntentIdAllocator()(repeated.document, intent);
    const plan = compileEditorIntent(repeated.document, {
      intent,
      mode: {
        type: "suggesting",
        revision: {
          id: ids.revisionId,
          author: "Latest reviewer",
          date: "2026-02-03T00:00:00Z",
        },
        newIds: ids.newIds,
      },
    }).unwrap();
    const joined = applyDocumentOps(repeated.document, plan.ops).unwrap();
    expect(
      applyDocumentOps(joined.document, [...joined.inverse, ...repeated.inverse]).unwrap().document,
    ).toStrictEqual(original);
    const saved = await createDocx(joined.document);
    const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
    const savedAgain = await createDocx(reopened);
    const reopenedAgain = await parseDocx(savedAgain, {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(reopened.package.document.content).toStrictEqual(
      joined.document.package.document.content,
    );
    expect(reopenedAgain.package.document.content).toStrictEqual(reopened.package.document.content);
    const modelParagraphs = (document: Document) =>
      container === "body"
        ? document.package.document.content
        : document.package.document.content.flatMap((block) =>
            block.type === "table"
              ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.content))
              : [],
          );
    const survivor = modelParagraphs(reopenedAgain).at(1);
    if (survivor?.type !== "paragraph") panic("Expected the pending join survivor");
    expect(survivor.propertyChanges).toEqual([
      {
        type: "paragraphPropertyChange",
        info: { id: 8, author: "Latest reviewer", date: "2026-02-03T00:00:00Z" },
        previousFormatting: { alignment: "start", styleId: "Normal" },
        currentFormatting: { alignment: "center", styleId: "Heading1" },
      },
    ]);
    expect(survivor.formatting).toEqual({ alignment: "center", styleId: "Heading1" });
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const resolution = {
        type: "resolveRevision",
        story: OP_STORIES.MAIN,
        revisionIds: [8],
        decision,
      } as const;
      const resolved = applyDocumentOps(reopenedAgain, [resolution]).unwrap();
      const resolvedBeforeSave = applyDocumentOps(joined.document, [resolution]).unwrap();
      expect(resolved.document.package.document.content).toStrictEqual(
        resolvedBeforeSave.document.package.document.content,
      );
      const reviewed = modelParagraphs(resolved.document).at(1);
      if (reviewed?.type !== "paragraph") panic("Expected the reviewed survivor");
      expect(reviewed.formatting).toEqual(
        decision === REVISION_DECISIONS.ACCEPT
          ? { alignment: "center", styleId: "Heading1" }
          : { alignment: "start", styleId: "Normal" },
      );
      expect(reviewed.propertyChanges).toBeUndefined();
      expect(applyDocumentOps(resolved.document, resolved.inverse).unwrap().document).toStrictEqual(
        reopenedAgain,
      );
    }
  },
);
