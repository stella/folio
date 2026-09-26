import { expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";
import { createEmptyDocx } from "../src/docx/rezip";
import { parseDocx } from "../src/docx/parser";
import { serializeDocument } from "../src/docx/serializer/documentSerializer";
import { fromProseDoc } from "../src/prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../src/prosemirror/conversion/toProseDoc";
import { ensureParaIdsInState } from "../src/prosemirror/extensions/features/ParaIdAllocatorExtension";
import { censusParagraphOrdinals } from "./selective-save-census-identity";

const wrap = (body: string) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}</w:body></w:document>`;
const paragraph = (id?: string) =>
  `<w:p${id ? ` w14:paraId="${id}"` : ""}><w:r><w:t>Editable target</w:t></w:r></w:p>`;

for (const coverage of ["authored", "none"] as const) {
  test(`maps the target after an extracted text box and non-paragraph blocks (${coverage})`, async () => {
    const box =
      '<w:p><w:r><w:pict><v:shape id="box" style="width:72pt;height:36pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Box text</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>';
    const source = wrap(
      '<w:bookmarkStart w:id="1" w:name="range"/><w:altChunk/>' +
        box +
        '<w:bookmarkEnd w:id="1"/>' +
        paragraph(coverage === "authored" ? "10000001" : undefined),
    );
    const zip = await JSZip.loadAsync(await createEmptyDocx());
    zip.file("word/document.xml", source);
    const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const state = ensureParaIdsInState(EditorState.create({ doc: toProseDoc(parsed) }));
    const directParagraphs: string[] = [];
    state.doc.forEach((node) => {
      if (node.type.name === "paragraph") directParagraphs.push(node.attrs["paraId"]);
    });
    // The text-box host is a source w:p but no longer a PM paragraph.
    expect(directParagraphs).toHaveLength(1);
    const targetId = directParagraphs.at(0);
    if (!targetId) panic("Expected allocated target ID");
    const serialized = serializeDocument(fromProseDoc(state.doc, parsed));
    expect(censusParagraphOrdinals(source, serialized).get(targetId)).toBe(1);
    expect(censusParagraphOrdinals(serialized, serialized).get(targetId)).toBe(1);
  });
}

test("authored IDs follow XML positions even when serialization expands earlier content", () => {
  const source = wrap(paragraph("10000001"));
  const serialized = wrap(paragraph("10000002") + paragraph("10000001"));
  expect(censusParagraphOrdinals(source, serialized).get("10000001")).toBe(0);
  expect(censusParagraphOrdinals(serialized, serialized).get("10000001")).toBe(1);
});

test("refuses ambiguous or unaligned minted identities", () => {
  expect(
    censusParagraphOrdinals(wrap(paragraph()), wrap(paragraph("10000001") + paragraph("10000002")))
      .size,
  ).toBe(0);
  expect(
    censusParagraphOrdinals(
      wrap(paragraph("10000001") + paragraph("10000001")),
      wrap(paragraph("10000001")),
    ).size,
  ).toBe(0);
  expect(
    censusParagraphOrdinals(wrap(paragraph()), wrap(paragraph("10000001") + paragraph("10000001")))
      .size,
  ).toBe(0);
  expect(
    censusParagraphOrdinals(wrap(paragraph("10000001")), wrap(paragraph("10000002"))).size,
  ).toBe(0);
});

test("raw selective saves can omit minted IDs without losing the exclusion", () => {
  const rawSaved = wrap(paragraph() + paragraph());
  const serialized = wrap(paragraph("10000001") + paragraph("10000002"));
  expect(censusParagraphOrdinals(rawSaved, serialized).get("10000002")).toBe(1);
});
