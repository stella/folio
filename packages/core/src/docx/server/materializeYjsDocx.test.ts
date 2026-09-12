import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";
import { initProseMirrorDoc, prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { expectTableCellAttrs } from "../../prosemirror/attrs";
import { schema } from "../../prosemirror/schema";
import { writeYjsParagraphSourceContract } from "../../prosemirror/yjsParagraphSourceContract";
import {
  ParagraphPropertySourceValidationError,
  PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR,
  TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR,
} from "../paragraphPropertySource";
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

const decodeCollaborativeDocument = (update: Uint8Array): EditorState["doc"] => {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, update);
  const document = initProseMirrorDoc(
    ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
    schema,
  ).doc;
  ydoc.destroy();
  return document;
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

const sourceWithCollapsedVerticalMergeProperties = async (): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (!documentXml) {
    throw new Error("Generated DOCX has no main document part");
  }
  zip.file(
    "word/document.xml",
    documentXml
      .replace("<w:document", '<w:document xmlns:x="urn:folio:test:paragraph-source"')
      .replace(
        /<w:body>[\s\S]*<\/w:body>/u,
        '<w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>' +
          '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>' +
          '<w:tr><w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Visible</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Top</w:t></w:r></w:p></w:tc></w:tr>' +
          '<w:tr><w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/><w:vMerge/></w:tcPr><w:p><w:pPr><w:ind w:left="1440"/><x:sentinel x:id="hidden-continuation"/></w:pPr></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Bottom</w:t></w:r></w:p></w:tc></w:tr>' +
          '</w:tbl><w:sectPr/></w:body>',
      ),
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

const collapsedContinuation = (document: EditorState["doc"]) => {
  let result:
    | { cellPos: number; cells: NonNullable<ReturnType<typeof expectTableCellAttrs>["_docxVMergeContinuationCells"]> }
    | undefined;
  document.descendants((node, pos) => {
    if (node.type.name !== "tableCell" && node.type.name !== "tableHeader") {
      return true;
    }
    const cells = expectTableCellAttrs(node)._docxVMergeContinuationCells;
    if (cells && cells.length > 0) {
      result = { cellPos: pos, cells };
    }
    return false;
  });
  if (!result) {
    throw new Error("Fixture did not collapse its vertical continuation cell");
  }
  return result;
};

const collapsedContinuationSourceToken = (paragraph: unknown): string => {
  if (typeof paragraph !== "object" || paragraph === null) {
    throw new Error("Fixture did not retain its hidden continuation paragraph");
  }
  const binding = Reflect.get(paragraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR);
  if (
    typeof binding !== "object" ||
    binding === null ||
    !("type" in binding) ||
    binding.type !== "source" ||
    !("token" in binding) ||
    typeof binding.token !== "string"
  ) {
    throw new Error("Fixture hidden continuation paragraph has no source binding");
  }
  return binding.token;
};

const HIDDEN_SOURCE_BINDING_CORRUPTIONS = [
  { expectedCode: "invalid_token", type: "absent" },
  { expectedCode: "invalid_token", type: "sourceWithoutToken" },
  { expectedCode: "invalid_token", type: "sourceWithNonStringToken" },
  { expectedCode: "invalid_token", type: "authoredWithToken" },
  { expectedCode: "unknown_token", type: "foreignFingerprint" },
  { expectedCode: "unknown_token", type: "unknownOrdinal" },
  { expectedCode: "duplicate_token", type: "duplicateVisibleToken" },
] as const;

type HiddenSourceBindingCorruption = (typeof HIDDEN_SOURCE_BINDING_CORRUPTIONS)[number]["type"];

type CorruptedHiddenBindingOptions = {
  corruption: HiddenSourceBindingCorruption;
  sourceToken: string;
  visibleToken: string;
};

const corruptedHiddenBinding = ({
  corruption,
  sourceToken,
  visibleToken,
}: CorruptedHiddenBindingOptions): unknown => {
  const separator = sourceToken.lastIndexOf(":");
  const ordinal = sourceToken.slice(separator + 1);
  const tokenPrefix = sourceToken.slice(0, separator);
  const fingerprintStart = tokenPrefix.indexOf(":") + 1;
  const fingerprint = tokenPrefix.slice(fingerprintStart);
  switch (corruption) {
    case "absent":
      return undefined;
    case "sourceWithoutToken":
      return { type: "source" };
    case "sourceWithNonStringToken":
      return { token: 1, type: "source" };
    case "authoredWithToken":
      return { token: sourceToken, type: "authored" };
    case "foreignFingerprint": {
      const firstCharacter = fingerprint.at(0);
      if (!firstCharacter) {
        throw new Error("Fixture source token has no fingerprint");
      }
      const foreignFingerprint = `${firstCharacter === "0" ? "1" : "0"}${fingerprint.slice(1)}`;
      return { token: `p1d:${foreignFingerprint}:${ordinal}`, type: "source" };
    }
    case "unknownOrdinal":
      return { token: `${tokenPrefix}:zz`, type: "source" };
    case "duplicateVisibleToken":
      return { token: visibleToken, type: "source" };
    default: {
      const exhaustiveCorruption: never = corruption;
      return exhaustiveCorruption;
    }
  }
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
  });

  test("rebinds a collapsed vertical-merge paragraph capture across repeated Yjs saves", async () => {
    const sourceDocx = await sourceWithCollapsedVerticalMergeProperties();
    const sourceDocument = await parseDocx(sourceDocx, { preloadFonts: false });
    const proseDocument = toProseDoc(sourceDocument);
    const { cells } = collapsedContinuation(proseDocument);
    const hiddenParagraph = cells.at(0)?.content.at(0);
    if (hiddenParagraph?.type !== "paragraph") {
      throw new Error("Fixture did not retain its hidden continuation paragraph");
    }
    const sourceToken = collapsedContinuationSourceToken(hiddenParagraph);

    const yjsUpdate = encodeCollaborativeDocument(proseDocument);
    const reconstructed = decodeCollaborativeDocument(yjsUpdate);
    const reconstructedParagraph = collapsedContinuation(reconstructed).cells.at(0)?.content.at(0);
    expect(collapsedContinuationSourceToken(reconstructedParagraph)).toBe(sourceToken);

    let output = await materializeYjsDocx({ sourceDocx, yjsUpdate });
    for (let repeatedSave = 0; repeatedSave < 2; repeatedSave += 1) {
      const repeatedSource = await parseDocx(output, { preloadFonts: false });
      output = await materializeYjsDocx({
        sourceDocx: output,
        yjsUpdate: encodeCollaborativeDocument(toProseDoc(repeatedSource)),
      });
    }
    const outputXml = await (await JSZip.loadAsync(output)).file("word/document.xml")?.async("text");

    expect(outputXml).toContain('<x:sentinel x:id="hidden-continuation"/>');
    expect(outputXml).toContain('<w:ind w:left="1440"/>');
    expect(outputXml).not.toContain(PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR);
    expect(outputXml).not.toContain(TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR);
  });

  test.each(HIDDEN_SOURCE_BINDING_CORRUPTIONS)(
    "rejects $type hidden continuation source binding",
    async ({ expectedCode, type }) => {
      const sourceDocx = await sourceWithCollapsedVerticalMergeProperties();
      const sourceDocument = await parseDocx(sourceDocx, { preloadFonts: false });
      const initialState = EditorState.create({ doc: toProseDoc(sourceDocument) });
      const { cellPos, cells } = collapsedContinuation(initialState.doc);
      const visibleToken = initialState.doc.child(0).attrs[PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR];
      const copiedCells = structuredClone(cells);
      const hiddenParagraph = copiedCells.at(0)?.content.at(0);
      if (typeof visibleToken !== "string" || hiddenParagraph?.type !== "paragraph") {
        throw new Error("Fixture did not expose both paragraph source identities");
      }
      const sourceToken = collapsedContinuationSourceToken(hiddenParagraph);
      const nextParagraph = { ...hiddenParagraph };
      const binding = corruptedHiddenBinding({ corruption: type, sourceToken, visibleToken });
      if (binding === undefined) {
        if (!Reflect.deleteProperty(nextParagraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR)) {
          throw new Error("Fixture could not remove the hidden source binding");
        }
      } else if (!Reflect.set(nextParagraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR, binding)) {
        throw new Error("Fixture could not replace the hidden source binding");
      }
      copiedCells[0]!.content[0] = nextParagraph;
      const corrupted = initialState.apply(
        initialState.tr.setNodeAttribute(cellPos, "_docxVMergeContinuationCells", copiedCells),
      );

      try {
        await materializeYjsDocx({
          sourceDocx,
          yjsUpdate: encodeCollaborativeDocument(corrupted.doc),
        });
        throw new Error("Expected hidden paragraph source validation to fail");
      } catch (error) {
        if (!(error instanceof FolioYjsDocxMaterializationError)) {
          throw error;
        }
        expect(error.code).toBe("source_mismatch");
        if (!(error.cause instanceof ParagraphPropertySourceValidationError)) {
          throw new Error("Expected a paragraph source validation cause");
        }
        expect(error.cause.code).toBe(expectedCode);
      }
    },
  );

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
    const duplicated = initialState.apply(
      initialState.tr.setNodeMarkup(secondPos, undefined, {
        ...second.attrs,
        [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: first.attrs[PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR],
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
    const sourceToken = first.attrs[PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR];
    if (typeof sourceToken !== "string") {
      throw new Error("Source fixture paragraph must carry a source token");
    }
    const unknownToken = `${sourceToken.slice(0, sourceToken.lastIndexOf(":"))}:zz`;
    const spoofed = initialState.apply(
      initialState.tr.setNodeMarkup(0, undefined, {
        ...first.attrs,
        [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: unknownToken,
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
