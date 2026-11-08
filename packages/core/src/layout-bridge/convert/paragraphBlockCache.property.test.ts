/**
 * Converting a document with the paragraph blocks kept from earlier
 * conversions equals converting every paragraph.
 *
 * Random runs of edits in the editor, typing, deleting across paragraphs,
 * splitting, list indent and outdent, undo and redo, and moving or copying a
 * whole paragraph elsewhere, over documents with lists, tables, notes,
 * comments, revisions, sections, fields, bookmarks, cross-references to list
 * numbers and anchored text boxes, are converted both ways after every edit,
 * with options that change along the way.
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
import type { Document } from "../../types/document";
import type { ToFlowBlocksOptions } from "./flowConversionShared";
import { ParagraphBlockCache } from "./paragraphBlockCache";
import { toFlowBlocks, WORD_UNSPECIFIED_FONT_SIZE } from "./toFlowBlocks";

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
    paragraph(text("Tail.")),
  ].join("");
  zip.set(
    "word/document.xml",
    xml.replace(/<w:body>.*<w:sectPr>/su, `<w:body>${body}<w:sectPr>`),
  );
  return zip.generate({ compression: "store" });
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
});

/** Counts the blocks it hands out, so the property cannot pass by never reusing one. */
class CountingCache extends ParagraphBlockCache {
  reused = 0;
  override take(...args: Parameters<ParagraphBlockCache["take"]>) {
    const block = super.take(...args);
    if (block) this.reused += 1;
    return block;
  }
}

describe("paragraph block cache", () => {
  test("converting with kept paragraph blocks equals converting every paragraph", async () => {
    const documents = await Promise.all(
      [...SHAPES.map(({ build }) => build), referencesPackage].map(
        async (build) => parseShapeDocument(await build()),
      ),
    );
    let reused = 0;
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
          const convert = () => {
            const constants = flowOptions(document, options);
            const kept = toFlowBlocks(view.state.doc, {
              ...constants,
              paragraphBlocks: cache,
            });
            expect(kept).toEqual(toFlowBlocks(view.state.doc, constants));
          };
          convert();
          for (const edit of edits) {
            applyEdit(view, edit, options);
            convert();
          }
          reused += cache.reused;
        },
      ),
      { id: "paragraph-block-cache-equals-full-conversion", numRuns: 60 },
    );
    expect(reused).toBeGreaterThan(0);
  });
});
