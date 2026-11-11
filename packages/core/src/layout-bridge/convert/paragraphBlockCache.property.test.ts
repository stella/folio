/**
 * Converting a document with the blocks kept from earlier conversions equals
 * converting every node.
 *
 * Random runs of edits in the editor, typing, deleting across paragraphs,
 * splitting, list indent and outdent, undo and redo, and moving or copying a
 * whole paragraph elsewhere, over documents with lists, tables, notes,
 * comments, revisions, sections, fields, bookmarks, cross-references to list
 * numbers, anchored text boxes, content controls, page breaks, run-in
 * headings and document grids, are converted both ways after every edit,
 * with options that change along the way. Every block handed out is frozen,
 * so a later conversion that wrote to a block it hands out again would fail.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { DOCUMENT_SHAPES } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
} from "../../__tests__/editorHarness";
import { blockIdsIn, renameBlockIds } from "../../__tests__/renameBlockIds";
import {
  hashParagraphBlock,
  serializeParagraphBlock,
} from "../../layout-engine/measure/cache";
import type { FlowBlock } from "../../layout-engine/types";
import type { Document } from "../../types/document";
import type { ToFlowBlocksOptions } from "./flowConversionShared";
import { ParagraphBlockCache } from "./paragraphBlockCache";
import { toFlowBlocks, WORD_UNSPECIFIED_FONT_SIZE } from "./toFlowBlocks";
import type { RecordedConversion } from "./topLevelConversion";

setDefaultTimeout(propertyTestTimeout(120_000));

const SHAPES = DOCUMENT_SHAPES.filter(({ id }) =>
  [
    "single-decimal-list",
    "mixed-lists",
    "style-numbered-headings",
    "outline-level-numbered",
    "tables",
    "notes",
    "comments",
    "invisible-markers",
    "tracked-changes",
    "pending-property-change",
    "sections",
    "fields-links-bookmarks",
    "rtl-cjk",
  ].includes(id),
);

const text = (value: string) =>
  `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const paragraph = (inner: string, numbered = false) =>
  `<w:p>${numbered ? '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' : ""}${inner}</w:p>`;
const bookmark = (id: number, inner: string) =>
  `<w:bookmarkStart w:id="${id}" w:name="_Ref${id}"/>${inner}<w:bookmarkEnd w:id="${id}"/>`;
/** A cross-reference showing the list number of the paragraph holding bookmark `id`. */
const numberReference = (id: number, shown: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref${id} \\r \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${text(shown)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
/** The same cross-reference, its number formatted apart from the space after it. */
const structuredNumberReference = (id: number, shown: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref${id} \\r \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>${shown}</w:t></w:r>${text(" ")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
/** A paragraph in a text frame beside the text. */
const framed = (inner: string) =>
  `<w:p><w:pPr><w:framePr w:w="2400" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:x="200" w:y="100"/></w:pPr>${inner}</w:p>`;
const anchoredTextBox = (id: number, value: string) =>
  `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2000000" cy="600000"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p>${text(value)}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

/** List items that other paragraphs refer to by number, and paragraphs anchoring text boxes. */
const referencesPackage = async (): Promise<Uint8Array> => {
  const list = DOCUMENT_SHAPES.find(({ id }) => id === "single-decimal-list");
  if (!list) throw new TypeError("The decimal list shape is missing.");
  const zip = await ZipArchive.read(await list.build());
  const xml = await zip.entry("word/document.xml")?.text();
  if (xml === undefined) throw new TypeError("The list shape has no body.");
  const body = [
    paragraph(text("Intro.")),
    paragraph(bookmark(1, text("Alpha")), true),
    paragraph(bookmark(2, text("Beta")), true),
    paragraph(
      text("See ") +
        numberReference(2, "2") +
        text(" and ") +
        anchoredTextBox(1, "Boxed"),
    ),
    paragraph(text("Gamma"), true),
    paragraph(text("Also ") + numberReference(1, "1") + text(" here")),
    paragraph(text("Delta ") + anchoredTextBox(2, "Second"), true),
    paragraph(
      text("Then ") + structuredNumberReference(2, "2") + text("again"),
    ),
    framed(text("Framed one")),
    framed(text("Framed two")),
    paragraph(text("Tail.")),
  ].join("");
  zip.set(
    "word/document.xml",
    xml.replace(/<w:body>.*<w:sectPr>/su, `<w:body>${body}<w:sectPr>`),
  );
  return zip.generate({ compression: "store" });
};

/**
 * Content controls, page breaks, run-in headings, grids, a closing table, a
 * cross-reference to a later list number and an empty section mark.
 */
const structuresPackage = async (): Promise<Uint8Array> => {
  const list = DOCUMENT_SHAPES.find(({ id }) => id === "single-decimal-list");
  if (!list) throw new TypeError("The decimal list shape is missing.");
  const zip = await ZipArchive.read(await list.build());
  const xml = await zip.entry("word/document.xml")?.text();
  if (xml === undefined) throw new TypeError("The list shape has no body.");
  const control = (tag: string, inner: string) =>
    `<w:sdt><w:sdtPr><w:alias w:val="Clause ${tag}"/><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
  const gridSection = (top: number, pitch: number) =>
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="${top}" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="${pitch}"/></w:sectPr>`;
  const body = [
    '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:p>',
    paragraph(text("Opening, see ") + numberReference(9, "4")),
    paragraph(text("First item"), true),
    control(
      "outer",
      paragraph(text("Controlled item"), true) +
        control("inner", paragraph(text("Nested control"))) +
        paragraph(text("Controlled tail")),
    ),
    paragraph(text("Second item"), true),
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
    `<w:p><w:pPr><w:rPr><w:specVanish/></w:rPr></w:pPr>${text("Run-in heading.")}</w:p>`,
    paragraph(text("Runs on after the heading.")),
    `<w:p><w:pPr>${gridSection(1000, 360)}</w:pPr>${text("Ends a gridded section.")}</w:p>`,
    paragraph(bookmark(9, text("Third item")), true),
    paragraph(text("After the grid.")),
    `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1200" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p>`,
    paragraph(text("Closing words.")),
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraph(text("Cell item"), true)}</w:tc></w:tr></w:tbl>`,
    "<w:p/>",
  ].join("");
  zip.set(
    "word/document.xml",
    xml.replace(
      /<w:body>.*<\/w:body>/su,
      `<w:body>${body}${gridSection(1440, 312)}</w:body>`,
    ),
  );
  return zip.generate({ compression: "store" });
};

/** Freeze every plain object and array a block holds. */
const deepFreeze = (value: unknown, seen = new Set<unknown>()): void => {
  if (typeof value !== "object" || value === null || seen.has(value)) return;
  const prototype: unknown = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== Array.prototype) return;
  seen.add(value);
  for (const nested of Object.values(value)) deepFreeze(nested, seen);
  Object.freeze(value);
};

const editArbitrary = fc.record({
  kind: fc.constantFrom(
    "type",
    "delete",
    "enter",
    "indent",
    "outdent",
    "undo",
    "redo",
    "move",
    "copy",
    "options",
  ),
  from: fc.double({ min: 0, max: 1, noNaN: true }),
  span: fc.double({ min: 0, max: 0.4, noNaN: true }),
  to: fc.double({ min: 0, max: 1, noNaN: true }),
  text: fc.constantFrom("x", "new words", " "),
});
type Edit = fc.RecordValue<typeof editArbitrary>;

/** Put the selection on the text nearest `from`, spanning `span` of the document. */
const select = (view: HeadlessEditorView, edit: Edit): void => {
  const { doc } = view.state;
  const size = doc.content.size;
  const from = Math.max(0, Math.min(size, Math.floor(edit.from * size)));
  const to = Math.min(size, from + Math.floor(edit.span * size));
  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.between(doc.resolve(from), doc.resolve(to)),
    ),
  );
};

/** Move or copy one top-level node, the same node, to another top-level gap. */
const relocate = (view: HeadlessEditorView, edit: Edit): void => {
  const { doc } = view.state;
  const count = doc.childCount;
  if (count < 2) return;
  const index = Math.min(count - 1, Math.floor(edit.from * count));
  const node = doc.child(index);
  let start = 0;
  for (let i = 0; i < index; i += 1) start += doc.child(i).nodeSize;
  const target = Math.min(count, Math.floor(edit.to * (count + 1)));
  let gap = 0;
  for (let i = 0; i < target; i += 1) gap += doc.child(i).nodeSize;
  const tr = view.state.tr.insert(gap, node);
  if (edit.kind === "move") {
    const moved = tr.mapping.map(start, 1);
    tr.delete(moved, moved + node.nodeSize);
  }
  view.dispatch(tr);
};

const applyEdit = (
  view: HeadlessEditorView,
  edit: Edit,
  options: ConstantChoices,
): void => {
  switch (edit.kind) {
    case "type":
      select(view, edit);
      view.typeText(edit.text);
      return;
    case "delete":
      select(view, edit);
      view.pressKey("Backspace");
      return;
    case "enter":
      select(view, { ...edit, span: 0 });
      view.pressKey("Enter");
      return;
    case "indent":
      select(view, edit);
      view.pressKey("Tab");
      return;
    case "outdent":
      select(view, edit);
      view.pressKey("Shift-Tab");
      return;
    case "undo":
      view.pressKey("Mod-z");
      return;
    case "redo":
      view.pressKey("Mod-Shift-z");
      return;
    case "move":
    case "copy":
      relocate(view, edit);
      return;
    case "options":
      if (edit.from < 0.5)
        options.showCommentHighlights = !options.showCommentHighlights;
      else options.alternate = options.alternate === "One" ? "Two" : "One";
      return;
    default: {
      const exhaustive: never = edit.kind;
      return exhaustive;
    }
  }
};

const EMBEDDED_FONTS: ReadonlyMap<string, string> = new Map([
  ["Body", "Body Embedded"],
]);

type ConstantChoices = { showCommentHighlights: boolean; alternate: string };

/** Options as a layout run builds them: some values afresh on every run. */
const flowOptions = (
  document: Document,
  { showCommentHighlights, alternate }: ConstantChoices,
): ToFlowBlocksOptions => ({
  pageContentHeight: 900,
  styles: document.package.styles,
  theme: document.package.theme ?? null,
  defaultSize: WORD_UNSPECIFIED_FONT_SIZE,
  defaultTabStopTwips: document.package.settings?.defaultTabStop ?? 720,
  showCommentHighlights,
  embeddedFontFamilyMap: EMBEDDED_FONTS,
  fontAlternates: new Map(
    ["calibri", "aptos", "times new roman", "arial", "cambria"].map(
      (family) => [family, `${alternate} ${family}`],
    ),
  ),
  automaticHyphenation: { enabled: true },
  finalSectionDocumentGridLinePitchTwips: 312,
  // Built afresh for every run, as a layout run builds them.
  justificationCompatibility: { type: "legacy" },
  tableIndentCompatibility: { type: "legacy" },
});

/** Every paragraph block, in table cells and text boxes too. */
const paragraphBlocksOf = (blocks: readonly FlowBlock[]) =>
  blocks.flatMap((block): FlowBlock[] => {
    switch (block.kind) {
      case "paragraph":
        return [block];
      case "table":
        return block.rows.flatMap((row) =>
          row.cells.flatMap((cell) => paragraphBlocksOf(cell.blocks)),
        );
      case "textBox":
        return paragraphBlocksOf(block.content);
      default:
        return [];
    }
  });

/** Counts the blocks it hands out, so the property cannot pass by never reusing one. */
class CountingCache extends ParagraphBlockCache {
  reused = 0;
  /** Each conversion's record, with the one it started from, oldest first. */
  readonly records: {
    readonly from: RecordedConversion | undefined;
    readonly record: RecordedConversion | undefined;
  }[] = [];
  private startedFrom: RecordedConversion | undefined;
  override take(...args: Parameters<ParagraphBlockCache["take"]>) {
    const block = super.take(...args);
    if (block) this.reused += 1;
    return block;
  }
  override lastConversion(
    ...args: Parameters<ParagraphBlockCache["lastConversion"]>
  ) {
    this.startedFrom = super.lastConversion(...args);
    return this.startedFrom;
  }
  override recordConversion(conversion: RecordedConversion | undefined) {
    this.records.push({ from: this.startedFrom, record: conversion });
    this.startedFrom = undefined;
    super.recordConversion(conversion);
  }
  /** The last conversion's record and the one it started from. */
  lastStep(): readonly [RecordedConversion, RecordedConversion] | undefined {
    const { from, record } = this.records.at(-1) ?? {};
    return from !== undefined && record !== undefined
      ? [from, record]
      : undefined;
  }
}

/**
 * The conversion with kept blocks equals the one from scratch up to the names
 * of its blocks, and names every block, row and cell once.
 */
function expectFreshConversion(
  kept: readonly FlowBlock[],
  fresh: readonly FlowBlock[],
): void {
  expect(renameBlockIds(kept)).toEqual(renameBlockIds(fresh));
  const ids = blockIdsIn(kept);
  expect(ids.length).toBeGreaterThan(0);
  expect(new Set(ids).size).toBe(ids.length);
}

/**
 * Every node a document still holds after an edit keeps the ids of its
 * blocks, whether its blocks are kept, moved or converted again, and whether
 * it stayed at either end or moved among the nodes the edit changed. A node
 * the document holds more than once keeps no particular one's ids. Returns
 * how many blocks kept their id, and how many of those belong to a node that
 * left its place among the changed nodes.
 */
function survivingNodesKeepTheirIds(
  before: RecordedConversion,
  after: RecordedConversion,
): { kept: number; movedInside: number } {
  const old = before.nodes;
  const now = after.nodes;
  let start = 0;
  while (
    start < old.length &&
    start < now.length &&
    old[start]?.node === now[start]?.node
  )
    start += 1;
  let end = 0;
  while (
    end < old.length - start &&
    end < now.length - start &&
    old[old.length - 1 - end]?.node === now[now.length - 1 - end]?.node
  )
    end += 1;
  type Nodes = RecordedConversion["nodes"];
  const placesOf = (nodes: Nodes) => {
    const places = new Map<Nodes[number]["node"], number[]>();
    nodes.forEach(({ node }, index) => {
      const at = places.get(node);
      if (at === undefined) places.set(node, [index]);
      else at.push(index);
    });
    return places;
  };
  const oldPlaces = placesOf(old);
  let kept = 0;
  let movedInside = 0;
  for (const [node, [to, ...again]] of placesOf(now)) {
    const [from, ...earlier] = oldPlaces.get(node) ?? [];
    if (to === undefined || from === undefined) continue;
    if (again.length > 0 || earlier.length > 0) continue;
    const was = old[from]?.blocks.map(({ id }) => id) ?? [];
    const is = now[to]?.blocks.map(({ id }) => id) ?? [];
    // A node converting again in other surroundings may give other blocks.
    if (was.length !== is.length) continue;
    expect(is).toEqual(was);
    kept += is.length;
    if (from !== to && to >= start && to < now.length - end)
      movedInside += is.length;
  }
  return { kept, movedInside };
}

describe("paragraph block cache", () => {
  test("converting with kept paragraph blocks equals converting every paragraph, and keeps their measurement keys", async () => {
    const documents = await Promise.all(
      [
        ...SHAPES.map(({ build }) => build),
        referencesPackage,
        structuresPackage,
      ].map(async (build) => parseShapeDocument(await build())),
    );
    let reused = 0;
    let handedOutAgain = 0;
    let keptIds = 0;
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 0, max: documents.length - 1 }),
        fc.boolean(),
        fc.array(editArbitrary, { minLength: 1, maxLength: 10 }),
        async (which, suggesting, edits) => {
          const document = documents[which]!;
          const view = new HeadlessEditorView(
            createHarnessState(document, suggesting ? "suggesting" : "editing"),
          );
          const cache = new CountingCache();
          const options = { showCommentHighlights: true, alternate: "One" };
          let last = new Set<FlowBlock>();
          const convert = () => {
            const constants = flowOptions(document, options);
            const kept = toFlowBlocks(view.state.doc, {
              ...constants,
              paragraphBlocks: cache,
            });
            expectFreshConversion(
              kept,
              toFlowBlocks(view.state.doc, constants),
            );
            for (const block of kept) if (last.has(block)) handedOutAgain += 1;
            last = new Set(kept);
            deepFreeze(kept);
            // The blocks each node projected, before the passes over the list.
            for (const node of cache.lastConversion(view.state.doc)?.nodes ??
              [])
              deepFreeze(node.blocks);
            // A moved paragraph's measurement key, taken from the block it
            // was moved from, is the key serializing it in full gives.
            for (const block of paragraphBlocksOf(kept))
              if (block.kind === "paragraph")
                expect(hashParagraphBlock(block)).toBe(
                  serializeParagraphBlock(block),
                );
          };
          convert();
          for (const edit of edits) {
            applyEdit(view, edit, options);
            convert();
            const step = cache.lastStep();
            if (step !== undefined)
              keptIds += survivingNodesKeepTheirIds(...step).kept;
          }
          reused += cache.reused;
        },
      ),
      { id: "paragraph-block-cache-equals-full-conversion", numRuns: 60 },
    );
    expect(reused).toBeGreaterThan(0);
    expect(handedOutAgain).toBeGreaterThan(0);
    expect(keptIds).toBeGreaterThan(0);
  });

  test("a cross-reference to a later list item shows the number the list gives now", async () => {
    const document = await parseShapeDocument(await structuresPackage());
    const view = new HeadlessEditorView(
      createHarnessState(document, "editing"),
    );
    const cache = new ParagraphBlockCache();
    const options = { showCommentHighlights: true, alternate: "One" };
    const convert = () => {
      const constants = flowOptions(document, options);
      const kept = toFlowBlocks(view.state.doc, {
        ...constants,
        paragraphBlocks: cache,
      });
      expectFreshConversion(kept, toFlowBlocks(view.state.doc, constants));
      return kept;
    };
    convert();
    // A new list item after the reference renumbers what it refers to; the
    // paragraph holding the reference is the same node, where it was.
    let firstItemEnd = 0;
    view.state.doc.descendants((node, position) => {
      if (node.isTextblock && node.textContent === "First item")
        firstItemEnd = position + node.nodeSize - 1;
      return firstItemEnd === 0;
    });
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, firstItemEnd),
      ),
    );
    view.pressKey("Enter");
    view.typeText("Inserted");
    convert();
  });

  test.each([
    { from: 0, past: 1 },
    { from: 0, past: 2 },
    { from: 2, past: -2 },
  ])(
    "a paragraph moved $past places among the nodes an edit changed keeps its ids, as do the nodes it passed",
    async ({ from, past }) => {
      const document = await parseShapeDocument(await referencesPackage());
      const view = new HeadlessEditorView(
        createHarnessState(document, "editing"),
      );
      const cache = new CountingCache();
      const options = { showCommentHighlights: true, alternate: "One" };
      const convert = () => {
        const constants = flowOptions(document, options);
        expectFreshConversion(
          toFlowBlocks(view.state.doc, {
            ...constants,
            paragraphBlocks: cache,
          }),
          toFlowBlocks(view.state.doc, constants),
        );
      };
      convert();
      // Move the node at `from` to the gap `past` nodes away, so the edit
      // reorders the nodes between the kept start and end. The nodes it
      // passes are the same nodes, now at other places in that range.
      const { doc } = view.state;
      const node = doc.child(from);
      const offsetOf = (index: number) => {
        let offset = 0;
        for (let i = 0; i < index; i += 1) offset += doc.child(i).nodeSize;
        return offset;
      };
      const start = offsetOf(from);
      const gap = offsetOf(past > 0 ? from + past + 1 : from + past);
      const tr = view.state.tr.insert(gap, node);
      const moved = tr.mapping.map(start, 1);
      view.dispatch(tr.delete(moved, moved + node.nodeSize));
      const index = from + past;
      const passed = Array.from({ length: Math.abs(past) }, (_, offset) =>
        doc.child(past > 0 ? from + 1 + offset : index + offset),
      );
      convert();
      const step = cache.lastStep();
      expect(step).toBeDefined();
      if (step === undefined) return;
      // Every node the move passed left its place and kept its ids; so did
      // the moved node, if the edit left it the same node.
      const { movedInside } = survivingNodesKeepTheirIds(...step);
      const survivors = new Set(step[0].nodes.map((at) => at.node));
      const left = step[1].nodes
        .slice(Math.min(from, index), Math.max(from, index) + 1)
        .filter((at) => survivors.has(at.node));
      expect(passed.every((one) => left.some((at) => at.node === one))).toBe(
        true,
      );
      expect(movedInside).toBe(
        left.reduce((count, { blocks }) => count + blocks.length, 0),
      );
      expect(movedInside).toBeGreaterThan(0);
    },
  );

  test("an edit that adds a block without moving the paragraphs after it keeps their ids", async () => {
    const document = await parseShapeDocument(await structuresPackage());
    const view = new HeadlessEditorView(
      createHarnessState(document, "editing"),
    );
    const cache = new CountingCache();
    const options = { showCommentHighlights: true, alternate: "One" };
    const convert = () => {
      const constants = flowOptions(document, options);
      const kept = toFlowBlocks(view.state.doc, {
        ...constants,
        paragraphBlocks: cache,
      });
      expectFreshConversion(kept, toFlowBlocks(view.state.doc, constants));
      return kept;
    };
    const before = convert();
    // A border on the empty section mark makes it paint a paragraph, so it
    // projects a block it did not; its size, and every position, stay.
    let markPos = -1;
    view.state.doc.forEach((node, position) => {
      if (
        node.type.name === "paragraph" &&
        node.content.size === 0 &&
        node.attrs["_sectionProperties"] != null
      )
        markPos = position;
    });
    expect(markPos).toBeGreaterThan(0);
    const mark = view.state.doc.nodeAt(markPos);
    view.dispatch(
      view.state.tr.setNodeMarkup(markPos, undefined, {
        ...mark?.attrs,
        borders: { bottom: { style: "single", size: 4 } },
      }),
    );
    const after = convert();
    expect(after.length).toBe(before.length + 1);
    // Every block but the one the mark now paints keeps its id.
    const afterIds = new Set(after.map(({ id }) => id));
    expect(before.filter(({ id }) => !afterIds.has(id))).toEqual([]);
    const step = cache.lastStep();
    expect(step).toBeDefined();
    if (step !== undefined)
      expect(survivingNodesKeepTheirIds(...step).kept).toBeGreaterThan(0);
  });

  test("paragraphs kept before a section mark take the line pitch it gives now", async () => {
    const document = await parseShapeDocument(await structuresPackage());
    const view = new HeadlessEditorView(
      createHarnessState(document, "editing"),
    );
    const cache = new ParagraphBlockCache();
    const options = { showCommentHighlights: true, alternate: "One" };
    const convert = () => {
      const constants = flowOptions(document, options);
      const kept = toFlowBlocks(view.state.doc, {
        ...constants,
        paragraphBlocks: cache,
      });
      expectFreshConversion(kept, toFlowBlocks(view.state.doc, constants));
      return kept;
    };
    convert();
    let markPos = -1;
    view.state.doc.forEach((node, position) => {
      const section = node.attrs["_sectionProperties"] as
        { docGrid?: { linePitch?: number } } | null | undefined;
      if (section?.docGrid?.linePitch === 360) markPos = position;
    });
    expect(markPos).toBeGreaterThan(0);
    const mark = view.state.doc.nodeAt(markPos);
    const section = mark?.attrs["_sectionProperties"] as {
      docGrid: Record<string, unknown>;
    };
    view.dispatch(
      view.state.tr.setNodeMarkup(markPos, undefined, {
        ...mark?.attrs,
        _sectionProperties: {
          ...section,
          docGrid: { ...section.docGrid, linePitch: 480 },
        },
      }),
    );
    convert();
  });

  test("a cross-reference formatted in parts shows the number the list gives now", async () => {
    const document = await parseShapeDocument(await referencesPackage());
    const view = new HeadlessEditorView(
      createHarnessState(document, "editing"),
    );
    const cache = new CountingCache();
    const options = { showCommentHighlights: true, alternate: "One" };
    const convert = () => {
      const constants = flowOptions(document, options);
      const kept = toFlowBlocks(view.state.doc, {
        ...constants,
        paragraphBlocks: cache,
      });
      expectFreshConversion(kept, toFlowBlocks(view.state.doc, constants));
    };
    convert();
    // A new list item before "Beta" renumbers it; the paragraphs referring to
    // it are unchanged nodes.
    let alphaEnd = 0;
    view.state.doc.descendants((node, position) => {
      if (node.isTextblock && node.textContent === "Alpha")
        alphaEnd = position + node.nodeSize - 1;
      return alphaEnd === 0;
    });
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, alphaEnd),
      ),
    );
    view.pressKey("Enter");
    view.typeText("Inserted");
    convert();
    expect(cache.reused).toBeGreaterThan(0);
  });
});
