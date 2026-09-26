import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import { Fragment, Slice } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  clearTrackedChanges,
  getChangedParagraphIds,
  hasStructuralChanges,
  hasUntrackedChanges,
  ParagraphChangeTrackerExtension,
} from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { ParaIdAllocatorExtension } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import type { Document } from "../types/document";
import { validateFolioDocumentModel } from "./modelValidation";
import { parseDocx } from "./parser";
import { attemptSelectiveSave } from "./selectiveSave";
import { findParagraphOffsets } from "./selectiveXmlPatch";

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PARA_ID_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const PARAGRAPH_TEXTS = ["Alpha beta", "Middle paragraph", "Last paragraph"] as const;
const PARAGRAPH_IDS = ["10000001", "10000002", "10000003"] as const;
const FINAL_SECTION = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

type ParagraphXmlOptions = { text: string; id: string };
const paragraphXml = ({ text, id }: ParagraphXmlOptions): string =>
  `<w:p w14:paraId="${id}" w:rsidR="0123ABCD"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const BODY = PARAGRAPH_TEXTS.map((text, index) => {
  const id = PARAGRAPH_IDS.at(index);
  if (!id) panic("Expected fixture paragraph identity");
  return paragraphXml({ text, id });
}).join("\n  ");
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' +
  paragraphXml({ text: "Cell paragraph", id: "20000001" }) +
  "</w:tc></w:tr></w:tbl>";
const documentXml = (body = BODY): string =>
  `${XML_DECLARATION}\n<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w14="${PARA_ID_NAMESPACE}">\n<w:body>\n  ${body}\n  ${FINAL_SECTION}\n</w:body>\n</w:document>`;

type PackageBytesOptions = { commentsXml?: string };
const packageBytes = async (
  source: string,
  options: PackageBytesOptions = {},
): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  zip.file("word/document.xml", source);
  zip.file(
    "customXml/item1.xml",
    '<opaque xmlns="urn:structural-save-test">  preserved &amp; untouched  </opaque>',
  );
  zip.file("word/media/opaque.bin", new Uint8Array([0, 255, 17, 128, 42]));
  if (options.commentsXml !== undefined) {
    zip.file("word/comments.xml", options.commentsXml);
    const contentTypes = zip.file("[Content_Types].xml");
    const relationships = zip.file("word/_rels/document.xml.rels");
    if (!contentTypes || !relationships) panic("Expected fixture package metadata");
    zip.file(
      "[Content_Types].xml",
      (await contentTypes.async("text")).replace(
        "</Types>",
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
      ),
    );
    zip.file(
      "word/_rels/document.xml.rels",
      (await relationships.async("text")).replace(
        "/>",
        '><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
      ),
    );
  }
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
};

const openEditor = async (buffer: ArrayBuffer) => {
  const base = await parseDocx(buffer, { preloadFonts: false });
  const doc = toProseDoc(base);
  const schema = doc.type.schema;
  const allocator = ParaIdAllocatorExtension().onSchemaReady({ schema });
  const tracker = ParagraphChangeTrackerExtension().onSchemaReady({ schema });
  const state = EditorState.create({
    doc,
    plugins: [...(allocator.plugins ?? []), ...(tracker.plugins ?? [])],
  });
  return { base, state };
};

const saveOptions = (state: EditorState) => ({
  changedParaIds: getChangedParagraphIds(state),
  structuralChange: hasStructuralChanges(state),
  hasUntrackedChanges: hasUntrackedChanges(state),
});

type SaveEditorOptions = { base: Document; state: EditorState; buffer: ArrayBuffer };
const saveEditor = ({ base, state, buffer }: SaveEditorOptions) =>
  attemptSelectiveSave(fromProseDoc(state.doc, base), buffer, saveOptions(state));

const firstBlockSize = (state: EditorState): number => {
  const first = state.doc.firstChild;
  if (!first) panic("Expected fixture block");
  return first.nodeSize;
};

const newParagraph = (state: EditorState, text: string): PMNode =>
  state.schema.node("paragraph", null, text ? state.schema.text(text) : undefined);

const bodyParagraphs = (doc: PMNode) => {
  const paragraphs: { node: PMNode; position: number; id: string }[] = [];
  doc.forEach((node, position) => {
    if (node.type.name !== "paragraph") return;
    const id = node.attrs["paraId"];
    if (typeof id !== "string") panic("Expected allocated paragraph identity");
    paragraphs.push({ node, position, id });
  });
  return paragraphs;
};

const readDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const entry = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!entry) panic("Expected word/document.xml");
  return entry.async("text");
};

type PreservedPartsOptions = { original: ArrayBuffer; saved: ArrayBuffer };
const expectOtherPartsUnchanged = async ({ original, saved }: PreservedPartsOptions) => {
  const before = await JSZip.loadAsync(original);
  const after = await JSZip.loadAsync(saved);
  expect(Object.keys(after.files).sort()).toEqual(Object.keys(before.files).sort());
  for (const [path, entry] of Object.entries(before.files)) {
    if (entry.dir || path === "word/document.xml") continue;
    const savedEntry = after.file(path);
    if (!savedEntry) panic(`Expected preserved part ${path}`);
    expect(await savedEntry.async("uint8array")).toEqual(await entry.async("uint8array"));
  }
};

const removeParagraphs = (xml: string, ids: ReadonlySet<string>): string => {
  const ranges = [...ids]
    .flatMap((id) => {
      const range = findParagraphOffsets(xml, id);
      return range ? [range] : [];
    })
    .sort((left, right) => right.start - left.start);
  let result = xml;
  for (const { start, end } of ranges) result = result.slice(0, start) + result.slice(end);
  return result;
};

type EditCase = {
  name: string;
  edit: (state: EditorState) => EditorState;
  expectedTexts: readonly string[];
};
const CASES = [
  {
    name: "Enter splits a paragraph",
    edit: (state) => state.apply(state.tr.split(6)),
    expectedTexts: ["Alpha", " beta", "Middle paragraph", "Last paragraph"],
  },
  ...PARAGRAPH_TEXTS.map((_, deletedIndex) => ({
    name: `delete paragraph ${deletedIndex + 1}`,
    edit: (state: EditorState) => {
      const target = bodyParagraphs(state.doc).at(deletedIndex);
      if (!target) panic("Expected paragraph to delete");
      return state.apply(state.tr.delete(target.position, target.position + target.node.nodeSize));
    },
    expectedTexts: PARAGRAPH_TEXTS.filter((_text, index) => index !== deletedIndex),
  })),
  {
    name: "delete all paragraphs to an empty document",
    edit: (state) => state.apply(state.tr.delete(0, state.doc.content.size)),
    expectedTexts: [""],
  },
  {
    name: "paste three paragraphs between surviving neighbours",
    edit: (state) =>
      state.apply(
        state.tr.replace(
          firstBlockSize(state),
          firstBlockSize(state),
          new Slice(
            Fragment.from(
              ["Paste one", "Paste two", "Paste three"].map((text) => newParagraph(state, text)),
            ),
            0,
            0,
          ),
        ),
      ),
    expectedTexts: [
      "Alpha beta",
      "Paste one",
      "Paste two",
      "Paste three",
      "Middle paragraph",
      "Last paragraph",
    ],
  },
  {
    name: "merge adjacent paragraphs",
    edit: (state) => state.apply(state.tr.join(firstBlockSize(state))),
    expectedTexts: ["Alpha betaMiddle paragraph", "Last paragraph"],
  },
  {
    name: "replace a paragraph without changing paragraph count",
    edit: (state) => {
      const target = bodyParagraphs(state.doc).at(1);
      if (!target) panic("Expected middle paragraph");
      return state.apply(
        state.tr.replaceWith(
          target.position,
          target.position + target.node.nodeSize,
          newParagraph(state, "Replacement"),
        ),
      );
    },
    expectedTexts: ["Alpha beta", "Replacement", "Last paragraph"],
  },
] satisfies readonly EditCase[];

describe("structural selective save through the editor pipeline", () => {
  for (const scenario of CASES) {
    test(scenario.name, async () => {
      const source = documentXml();
      const buffer = await packageBytes(source);
      const { base, state } = await openEditor(buffer);
      const edited = scenario.edit(state);
      expect(hasStructuralChanges(edited)).toBe(true);
      expect(hasUntrackedChanges(edited)).toBe(false);
      const expectedParagraphs = bodyParagraphs(edited.doc);
      const expectedIds = expectedParagraphs.map(({ id }) => id);
      expect(new Set(expectedIds).size).toBe(expectedIds.length);
      for (const id of expectedIds) expect(id).toMatch(/^[0-9A-F]{8}$/u);
      const saved = await saveEditor({ base, state: edited, buffer });
      if (!saved) panic(`Expected selective save for ${scenario.name}`);
      const reloaded = await openEditor(saved);
      const actualParagraphs = bodyParagraphs(reloaded.state.doc);
      expect(actualParagraphs.map(({ node }) => node.textContent)).toEqual(scenario.expectedTexts);
      expect(actualParagraphs.map(({ id }) => id)).toEqual(expectedIds);
      const xml = await readDocumentXml(saved);
      const originalParagraphs = bodyParagraphs(state.doc);
      const untouched = new Set(
        originalParagraphs
          .filter(({ node, id }) =>
            expectedParagraphs.some((current) => current.id === id && current.node.eq(node)),
          )
          .map(({ id }) => id),
      );
      for (const id of untouched) {
        const before = findParagraphOffsets(source, id);
        const after = findParagraphOffsets(xml, id);
        if (!before || !after) panic("Expected untouched paragraph in source and result");
        expect(xml.slice(after.start, after.end)).toBe(source.slice(before.start, before.end));
      }
      // Removing the edited/inserted/deleted spans exposes exactly the same
      // source scaffold: whitespace, body attributes and final section bytes.
      expect(removeParagraphs(xml, new Set(expectedIds.filter((id) => !untouched.has(id))))).toBe(
        removeParagraphs(
          source,
          new Set(originalParagraphs.filter(({ id }) => !untouched.has(id)).map(({ id }) => id)),
        ),
      );
      await expectOtherPartsUnchanged({ original: buffer, saved });
    });
  }

  test("repeated save after clearing the tracker preserves newly inserted identities", async () => {
    const buffer = await packageBytes(documentXml());
    const { base, state } = await openEditor(buffer);
    const split = state.apply(state.tr.split(6));
    const first = await saveEditor({ base, state: split, buffer });
    if (!first) panic("Expected split save");
    const cleared = split.apply(clearTrackedChanges(split));
    const target = bodyParagraphs(cleared.doc).at(1);
    if (!target) panic("Expected inserted paragraph");
    const typed = cleared.apply(cleared.tr.insertText("!", target.position + 1));
    const second = await saveEditor({ base, state: typed, buffer: first });
    if (!second) panic("Expected edit of previously inserted paragraph to save selectively");
    const reloaded = await openEditor(second);
    expect(bodyParagraphs(reloaded.state.doc).map(({ node }) => node.textContent)).toEqual([
      "Alpha",
      "! beta",
      "Middle paragraph",
      "Last paragraph",
    ]);
    expect(bodyParagraphs(reloaded.state.doc).map(({ id }) => id)).toEqual(
      bodyParagraphs(typed.doc).map(({ id }) => id),
    );
    await expectOtherPartsUnchanged({ original: first, saved: second });
  });

  test("an unchanged table remains a byte-exact anchor", async () => {
    const source = documentXml(
      paragraphXml({ text: "Alpha beta", id: "10000001" }) +
        TABLE +
        paragraphXml({ text: "After table", id: "10000002" }),
    );
    const buffer = await packageBytes(source);
    const { base, state } = await openEditor(buffer);
    const edited = state.apply(state.tr.split(6));
    const saved = await saveEditor({ base, state: edited, buffer });
    if (!saved) panic("Expected split beside unchanged table to save selectively");
    expect(await readDocumentXml(saved)).toContain(TABLE);
    await expectOtherPartsUnchanged({ original: buffer, saved });
  });

  test("noncanonical namespace aliases remain bound after inserting a paragraph", async () => {
    const source = documentXml()
      .replaceAll("w:", "word:")
      .replaceAll("xmlns:w=", "xmlns:word=")
      .replaceAll("w14:", "identity:")
      .replaceAll("xmlns:w14=", "xmlns:identity=");
    const buffer = await packageBytes(source);
    const { base, state } = await openEditor(buffer);
    const edited = state.apply(state.tr.split(6));
    const saved = await saveEditor({ base, state: edited, buffer });
    if (!saved) panic("Expected aliased source to save selectively");
    const reloaded = await openEditor(saved);
    expect(bodyParagraphs(reloaded.state.doc).map(({ node }) => node.textContent)).toEqual([
      "Alpha",
      " beta",
      "Middle paragraph",
      "Last paragraph",
    ]);
    expect(await readDocumentXml(saved)).toContain('<word:p identity:paraId="10000003"');
    await expectOtherPartsUnchanged({ original: buffer, saved });
  });
});

describe("structural save fallback boundaries", () => {
  test("cell paragraph insertion falls back", async () => {
    const buffer = await packageBytes(documentXml(BODY + TABLE));
    const { base, state } = await openEditor(buffer);
    let cellParagraphPosition: number | undefined;
    state.doc.descendants((node, position) => {
      if (node.type.name === "paragraph" && node.attrs["paraId"] === "20000001") {
        cellParagraphPosition = position;
      }
    });
    if (cellParagraphPosition === undefined) panic("Expected cell paragraph");
    const edited = state.apply(state.tr.split(cellParagraphPosition + 5));
    expect(await saveEditor({ base, state: edited, buffer })).toBeNull();
  });

  test("table row insertion falls back", async () => {
    const buffer = await packageBytes(documentXml(BODY + TABLE));
    const { base, state } = await openEditor(buffer);
    let row: { node: PMNode; position: number } | undefined;
    state.doc.descendants((node, position) => {
      if (node.type.name === "tableRow") row = { node, position };
    });
    if (!row) panic("Expected table row");
    const edited = state.apply(state.tr.insert(row.position + row.node.nodeSize, row.node));
    expect(await saveEditor({ base, state: edited, buffer })).toBeNull();
  });

  test("paragraph section endpoint deletion falls back", async () => {
    const sectionParagraph = paragraphXml({ text: "Section end", id: "10000004" }).replace(
      "<w:r>",
      '<w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r>',
    );
    const buffer = await packageBytes(documentXml(sectionParagraph + BODY));
    const { base, state } = await openEditor(buffer);
    const edited = state.apply(state.tr.delete(0, firstBlockSize(state)));
    expect(await saveEditor({ base, state: edited, buffer })).toBeNull();
  });

  test("comment range crossing a split falls back", async () => {
    const body =
      paragraphXml({ text: "Alpha beta", id: "10000001" }).replace(
        "<w:r>",
        '<w:commentRangeStart w:id="7"/><w:r>',
      ) +
      paragraphXml({ text: "Range end", id: "10000002" }).replace(
        "</w:p>",
        '<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>',
      );
    const buffer = await packageBytes(documentXml(body), {
      commentsXml: `${XML_DECLARATION}<w:comments xmlns:w="${WORD_NAMESPACE}" xmlns:w14="${PARA_ID_NAMESPACE}"><w:comment w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">${paragraphXml({ text: "Comment thread", id: "30000001" })}</w:comment></w:comments>`,
    });
    const { base, state } = await openEditor(buffer);
    expect(base.package.document.comments).toHaveLength(1);
    const edited = state.apply(state.tr.split(6));
    expect(await saveEditor({ base, state: edited, buffer })).toBeNull();
  });

  test("new numbering definitions fall back", async () => {
    const buffer = await packageBytes(documentXml());
    const { base, state } = await openEditor(buffer);
    const edited = state.apply(state.tr.split(6));
    const model = fromProseDoc(edited.doc, base);
    model.package.numbering = {
      abstractNums: [
        { abstractNumId: 0, levels: [{ ilvl: 0, start: 1, numFmt: "decimal", lvlText: "%1." }] },
      ],
      nums: [{ numId: 1, abstractNumId: 0 }],
    };
    const first = model.package.document.content.at(0);
    if (!first || first.type !== "paragraph") panic("Expected numbered paragraph");
    first.formatting = { ...first.formatting, numPr: { kind: "reference", numId: 1, ilvl: 0 } };
    expect(validateFolioDocumentModel(model).valid).toBe(true);
    expect(await attemptSelectiveSave(model, buffer, saveOptions(edited))).toBeNull();
  });

  test("source paragraphs without authored identities fall back", async () => {
    const buffer = await packageBytes(documentXml().replaceAll(/ w14:paraId="[0-9A-F]{8}"/gu, ""));
    const { base, state } = await openEditor(buffer);
    const edited = state.apply(state.tr.split(6));
    expect(await saveEditor({ base, state: edited, buffer })).toBeNull();
  });
});
