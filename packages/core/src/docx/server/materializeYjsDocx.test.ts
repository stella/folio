import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";
import { prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { writeYjsParagraphSourceContract } from "../../prosemirror/yjsParagraphSourceContract";
import { parseDocx } from "../parser";
import { createDocx, createEmptyDocx } from "../rezip";
import { createEmptyDocument } from "../../utils/createDocument";
import { extractDocxText } from "./extractDocxText";
import {
  FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME,
  FOLIO_YJS_UPDATE_MAX_BYTES,
  FolioYjsDocxMaterializationError,
  materializeYjsDocx,
} from "./materializeYjsDocx";

const encodeCollaborativeDocument = (document: EditorState["doc"]): Uint8Array => {
  const ydoc = new Y.Doc();
  prosemirrorToYXmlFragment(document, ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME));
  writeYjsParagraphSourceContract(ydoc, document);
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

const createCollaborativeUpdate = async (sourceDocx: ArrayBuffer, text: string) => {
  const sourceDocument = await parseDocx(sourceDocx, { preloadFonts: false });
  const initialState = EditorState.create({ doc: toProseDoc(sourceDocument) });
  const bodyEnd = initialState.doc.content.size - 1;
  const nextState = initialState.apply(
    initialState.tr.replaceWith(1, bodyEnd, initialState.schema.text(text)),
  );
  return encodeCollaborativeDocument(nextState.doc);
};

const sourceWithDistinctParagraphProperties = async (): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (!documentXml) {
    throw new Error("Generated DOCX has no main document part");
  }
  const paragraph = (text: string, sentinel: string, left: number, paraId?: string) =>
    `<w:p${paraId ? ` w14:paraId="${paraId}"` : ""}><w:pPr><w:ind w:left="${String(left)}"/><x:sentinel x:id="${sentinel}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  zip.file(
    "word/document.xml",
    documentXml
      .replace("<w:document", '<w:document xmlns:x="urn:folio:test:paragraph-source"')
      .replace(
        /<w:body>[\s\S]*<\/w:body>/u,
        `<w:body>${paragraph("Idless", "idless", 720)}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="7200"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/></w:tcPr>${paragraph("Duplicate A", "duplicate-a", 1440, "12345678")}</w:tc></w:tr></w:tbl>${paragraph("Duplicate B", "duplicate-b", 2160, "12345678")}<w:sectPr/></w:body>`,
      ),
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

describe("materializeYjsDocx", () => {
  test("replaces body content while preserving the source DOCX package", async () => {
    const sourceDocx = await createDocx(createEmptyDocument({ initialText: "Original text" }));
    const yjsUpdate = await createCollaborativeUpdate(sourceDocx, "Collaborative text");

    const output = await materializeYjsDocx({ sourceDocx, yjsUpdate });
    const extracted = await extractDocxText(output);
    const [sourceZip, outputZip] = await Promise.all([
      JSZip.loadAsync(sourceDocx),
      JSZip.loadAsync(output),
    ]);

    expect(extracted.paragraphs.map(({ text }) => text)).toContain("Collaborative text");
    expect(extracted.paragraphs.map(({ text }) => text)).not.toContain("Original text");
    expect(await outputZip.file("word/styles.xml")?.async("text")).toBe(
      await sourceZip.file("word/styles.xml")?.async("text"),
    );
  });

  test("rebinds id-less and duplicate-id paragraph properties after Yjs reconstruction", async () => {
    const sourceDocx = await sourceWithDistinctParagraphProperties();
    const sourceDocument = await parseDocx(sourceDocx, { preloadFonts: false });
    const initialState = EditorState.create({ doc: toProseDoc(sourceDocument) });
    const paragraphPositions: number[] = [];
    initialState.doc.descendants((node, pos) => {
      if (node.type.name === "paragraph") {
        paragraphPositions.push(pos);
        return false;
      }
      return true;
    });
    const transaction = initialState.tr;
    for (const pos of paragraphPositions.toReversed()) {
      transaction.insertText("!", pos + 1);
    }
    const yjsUpdate = encodeCollaborativeDocument(initialState.apply(transaction).doc);

    const output = await materializeYjsDocx({ sourceDocx, yjsUpdate });
    const outputXml = await (
      await JSZip.loadAsync(output)
    )
      .file("word/document.xml")
      ?.async("text");

    expect(outputXml).toContain('<x:sentinel x:id="idless"/>');
    expect(outputXml).toContain('<x:sentinel x:id="duplicate-a"/>');
    expect(outputXml).toContain('<x:sentinel x:id="duplicate-b"/>');
    expect(outputXml).toContain('<w:ind w:left="720"/>');
    expect(outputXml).toContain('<w:ind w:left="1440"/>');
    expect(outputXml).toContain('<w:ind w:left="2160"/>');
    expect(outputXml).not.toContain("_docxParagraphSource");
    expect(outputXml).not.toContain("folio-ppr-v1");
    expect(outputXml).not.toContain("p1d:");
    expect(outputXml).not.toContain("folio-ppr-v3");
    expect(outputXml).not.toContain("p3s:");
  });

  test("rejects a state update without Folio's document fragment", async () => {
    const sourceDocx = await createDocx(createEmptyDocument());
    const ydoc = new Y.Doc();
    ydoc.getMap("unrelated").set("value", true);
    const yjsUpdate = Y.encodeStateAsUpdate(ydoc);
    ydoc.destroy();

    await expect(materializeYjsDocx({ sourceDocx, yjsUpdate })).rejects.toMatchObject({
      _tag: "FolioYjsDocxMaterializationError",
      code: "missing_document",
    } satisfies Partial<FolioYjsDocxMaterializationError>);
  });

  test("rejects collaboration state without its source contract", async () => {
    const sourceDocx = await createDocx(createEmptyDocument({ initialText: "Original" }));
    const sourceDocument = await parseDocx(sourceDocx, { preloadFonts: false });
    const ydoc = new Y.Doc();
    prosemirrorToYXmlFragment(
      toProseDoc(sourceDocument),
      ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
    );
    const yjsUpdate = Y.encodeStateAsUpdate(ydoc);
    ydoc.destroy();

    await expect(materializeYjsDocx({ sourceDocx, yjsUpdate })).rejects.toMatchObject({
      _tag: "FolioYjsDocxMaterializationError",
      code: "source_mismatch",
    } satisfies Partial<FolioYjsDocxMaterializationError>);
  });

  test("rejects collaboration state paired with a different source package", async () => {
    const sourceDocx = await createDocx(createEmptyDocument({ initialText: "Source A" }));
    const otherSourceDocx = await createDocx(createEmptyDocument({ initialText: "Source B" }));
    const yjsUpdate = await createCollaborativeUpdate(sourceDocx, "Edited A");

    await expect(
      materializeYjsDocx({ sourceDocx: otherSourceDocx, yjsUpdate }),
    ).rejects.toMatchObject({
      _tag: "FolioYjsDocxMaterializationError",
      code: "source_mismatch",
    } satisfies Partial<FolioYjsDocxMaterializationError>);
  });

  test("rejects duplicated paragraph-property source tokens", async () => {
    const sourceDocx = await sourceWithDistinctParagraphProperties();
    const sourceDocument = await parseDocx(sourceDocx, { preloadFonts: false });
    const initialState = EditorState.create({ doc: toProseDoc(sourceDocument) });
    const paragraphPositions: number[] = [];
    initialState.doc.descendants((node, pos) => {
      if (node.type.name === "paragraph") {
        paragraphPositions.push(pos);
        return false;
      }
      return true;
    });
    const firstPos = paragraphPositions.at(0);
    const secondPos = paragraphPositions.at(1);
    if (firstPos === undefined || secondPos === undefined) {
      throw new Error("Source fixture must contain two paragraphs");
    }
    const first = initialState.doc.nodeAt(firstPos);
    const second = initialState.doc.nodeAt(secondPos);
    if (!first || !second) {
      throw new Error("Source fixture lost a paragraph");
    }
    const firstState = first.attrs["_paragraphPropertyState"];
    const secondState = second.attrs["_paragraphPropertyState"];
    if (
      typeof firstState !== "object" ||
      firstState === null ||
      !("type" in firstState) ||
      firstState.type !== "imported" ||
      !("token" in firstState) ||
      typeof firstState.token !== "string" ||
      typeof secondState !== "object" ||
      secondState === null ||
      !("type" in secondState) ||
      secondState.type !== "imported"
    ) {
      throw new Error("Source fixture paragraphs must carry imported property state");
    }
    const duplicated = initialState.apply(
      initialState.tr.setNodeMarkup(secondPos, undefined, {
        ...second.attrs,
        _paragraphPropertyState: { ...secondState, token: firstState.token },
      }),
    );
    const yjsUpdate = encodeCollaborativeDocument(duplicated.doc);

    await expect(materializeYjsDocx({ sourceDocx, yjsUpdate })).rejects.toMatchObject({
      _tag: "FolioYjsDocxMaterializationError",
      code: "source_mismatch",
    } satisfies Partial<FolioYjsDocxMaterializationError>);
  });

  test("rejects a well-formed token absent from the exact source", async () => {
    const sourceDocx = await sourceWithDistinctParagraphProperties();
    const sourceDocument = await parseDocx(sourceDocx, { preloadFonts: false });
    const initialState = EditorState.create({ doc: toProseDoc(sourceDocument) });
    const first = initialState.doc.child(0);
    const firstState = first.attrs["_paragraphPropertyState"];
    if (
      typeof firstState !== "object" ||
      firstState === null ||
      !("type" in firstState) ||
      firstState.type !== "imported" ||
      !("token" in firstState) ||
      typeof firstState.token !== "string"
    ) {
      throw new Error("Source fixture paragraph must carry a source token");
    }
    const sourceToken = firstState.token;
    const unknownToken = `${sourceToken.slice(0, sourceToken.lastIndexOf(":"))}:zz`;
    const spoofed = initialState.apply(
      initialState.tr.setNodeMarkup(0, undefined, {
        ...first.attrs,
        _paragraphPropertyState: { ...firstState, token: unknownToken },
      }),
    );
    const yjsUpdate = encodeCollaborativeDocument(spoofed.doc);

    await expect(materializeYjsDocx({ sourceDocx, yjsUpdate })).rejects.toMatchObject({
      _tag: "FolioYjsDocxMaterializationError",
      code: "source_mismatch",
    } satisfies Partial<FolioYjsDocxMaterializationError>);
  });

  test("rejects an oversized update before parsing the source DOCX", async () => {
    await expect(
      materializeYjsDocx({
        sourceDocx: new Uint8Array([1, 2, 3]),
        yjsUpdate: new Uint8Array(FOLIO_YJS_UPDATE_MAX_BYTES + 1),
      }),
    ).rejects.toMatchObject({
      _tag: "FolioYjsDocxMaterializationError",
      code: "update_too_large",
    } satisfies Partial<FolioYjsDocxMaterializationError>);
  });
});
