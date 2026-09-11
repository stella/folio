import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import { Fragment, Slice } from "prosemirror-model";

import { tableFromTemplate } from "../ai-edits/table-template";
import type { Document, Paragraph } from "../types/document";
import { expectTableCellAttrs } from "../prosemirror/attrs";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ensureParaIdsInDoc } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { replaceTextInDocument } from "../utils/replaceText";
import { withoutOrphanCommentRanges } from "./commentRangeIntegrity";
import { parseDocx } from "./parser";
import { DATE_UTC_NAMESPACE_URI } from "./trackedChangeInfo";
import {
  ParagraphPropertySourceValidationError,
  assignParagraphPropertySource,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
} from "./paragraphPropertySource";
import { createEmptyDocx, repackDocx } from "./rezip";
import { consolidateParagraph } from "./runConsolidator";
import { serializeParagraph, serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { unzipDocx } from "./unzip";
import { isSafeCapturedXmlDocument } from "./verbatimCapture";
import {
  getChildElements,
  getAttributeByNamespaceUri,
  getLocalName,
  NAMESPACES,
  parseXmlDocument,
  type XmlElement,
} from "./xmlParser";

const SOURCE_PROPERTIES = `<w:pPr>
  <w:ind w:left="720" w:leftChars="100"/>
  <w:cnfStyle w:val="000000100000" w:oddHBand="1"/>
  <w:rPr><w:bCs/><w:sz w:val="21"/><w:szCs w:val="22"/><w:noProof/></w:rPr>
</w:pPr>`;

const STRICT_WORDPROCESSINGML_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";

const adversarialParagraphPropertiesXml = (
  name: string,
  sourceXml: string,
  before: string,
  after: string,
): string => {
  switch (name) {
    case "unclosed root":
      return '<w:pPr><w:ind w:left="720"/><w:replay-unclosed>';
    case "mismatched root":
      return "<w:pPr></w:rPr>";
    case "named entity":
      return '<w:pPr><w:ind w:left="&custom;"/></w:pPr>';
    case "invalid numeric entity":
      return '<w:pPr><w:ind w:left="&#0;"/></w:pPr>';
    default:
      return `${before}${sourceXml}${after}`;
  }
};

const documentWithSourceProperties = async (
  paragraphProperties = SOURCE_PROPERTIES,
  paraId: string | null = "12345678",
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const documentFile = zip.file("word/document.xml");
  const documentXml = await documentFile?.async("text");
  if (!documentXml) {
    panic("The generated package has no main document part.");
  }
  zip.file(
    "word/document.xml",
    documentXml.replace(
      /<w:body>[\s\S]*<\/w:body>/u,
      `<w:body><w:p${paraId ? ` w14:paraId="${paraId}"` : ""}>${paragraphProperties}<w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body>`,
    ),
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

const documentWithInheritedStrictProperties = async (): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const documentFile = zip.file("word/document.xml");
  const documentXml = await documentFile?.async("text");
  if (!documentXml) {
    panic("The generated package has no main document part.");
  }
  zip.file(
    "word/document.xml",
    documentXml
      .replaceAll(NAMESPACES.w, STRICT_WORDPROCESSINGML_NAMESPACE)
      .replace(
        /<w:body>[\s\S]*<\/w:body>/u,
        '<w:body><w:p w14:paraId="12345678"><w:pPr><w:spacing w:before="12pt"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body>',
      ),
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

const documentWithDuplicateParagraphIds = async (): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const documentFile = zip.file("word/document.xml");
  const documentXml = await documentFile?.async("text");
  if (!documentXml) {
    panic("The generated package has no main document part.");
  }
  const paragraph = (text: string, left: string) =>
    `<w:p w14:paraId="12345678">${SOURCE_PROPERTIES.replace('w:left="720"', `w:left="${left}"`)}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  zip.file(
    "word/document.xml",
    documentXml.replace(
      /<w:body>[\s\S]*<\/w:body>/u,
      `<w:body>${paragraph("First", "720")}${paragraph("Second", "1440")}<w:sectPr/></w:body>`,
    ),
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

const canonicalElement = (element: XmlElement): string => {
  const attributes = Object.entries(element.attributes ?? {})
    .toSorted(([left], [right]) => left.localeCompare(right))
    .map(([name, value]) => `${name}=${String(value)}`)
    .join(" ");
  return `<${getLocalName(element.name)} ${attributes}>${getChildElements(element)
    .map(canonicalElement)
    .join("")}</>`;
};

const firstParagraphProperties = async (buffer: ArrayBuffer): Promise<XmlElement> => {
  const { documentXml } = await unzipDocx(buffer);
  if (typeof documentXml !== "string") {
    panic("The package has no main document part.");
  }
  const root = parseXmlDocument(documentXml);
  if (!root) {
    panic("The main document part could not be parsed.");
  }
  for (const element of getChildElements(root)) {
    const descendants = [element];
    while (descendants.length > 0) {
      const current = descendants.shift();
      if (!current) {
        break;
      }
      if (getLocalName(current.name) === "pPr") {
        return current;
      }
      descendants.push(...getChildElements(current));
    }
  }
  panic("The document has no paragraph properties.");
};

const firstParagraph = (document: Document): Paragraph => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    panic("The parsed document has no first paragraph.");
  }
  return block;
};

describe("paragraph properties survive a no-edit full repack", () => {
  test.each([
    ["typed model", (document: Document): Document => document],
    [
      "editable model",
      (document: Document): Document => fromProseDoc(toProseDoc(document), document),
    ],
  ] as const)("through the %s", async (_name, roundTrip) => {
    const source = await documentWithSourceProperties();
    const before = canonicalElement(await firstParagraphProperties(source));
    const parsed = await parseDocx(source, { preloadFonts: false });
    const saved = await repackDocx(roundTrip(parsed), { updateModifiedDate: false });

    expect(canonicalElement(await firstParagraphProperties(saved))).toBe(before);
  });

  test("the editable model recovers the source by paragraph identity", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const proseDoc = toProseDoc(parsed);

    expect(JSON.stringify(proseDoc.toJSON())).not.toContain("propertySource");
    expect(getParagraphPropertySource(firstParagraph(fromProseDoc(proseDoc, parsed)))).toEqual(
      getParagraphPropertySource(firstParagraph(parsed)),
    );
  });

  test("an unchanged id-less editable paragraph keeps its private source identity", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(SOURCE_PROPERTIES, null), {
      preloadFonts: false,
    });
    const restored = fromProseDoc(toProseDoc(parsed), parsed);

    expect(getParagraphPropertySource(firstParagraph(restored))).toEqual(
      getParagraphPropertySource(firstParagraph(parsed)),
    );
  });

  test("orphan-comment cleanup preserves the parser owner through editable replay", async () => {
    const source = await documentWithSourceProperties();
    const before = canonicalElement(await firstParagraphProperties(source));
    const parsed = await parseDocx(source, { preloadFonts: false });
    const sourceParagraph = firstParagraph(parsed);
    sourceParagraph.content.unshift({ type: "commentRangeStart", id: 999 });

    const restored = fromProseDoc(toProseDoc(parsed), parsed);
    const restoredParagraph = firstParagraph(restored);
    expect(getParagraphPropertySource(restoredParagraph)).toEqual(
      getParagraphPropertySource(sourceParagraph),
    );

    const cleaned = withoutOrphanCommentRanges(restored);
    const cleanedParagraph = firstParagraph(cleaned);
    expect(cleanedParagraph).not.toBe(restoredParagraph);
    expect(getParagraphPropertySource(cleanedParagraph)).toEqual(
      getParagraphPropertySource(sourceParagraph),
    );
    expect(getParagraphPropertySource(cleanedParagraph)).toEqual(
      getParagraphPropertySource(sourceParagraph),
    );
    expect(
      canonicalElement(
        await firstParagraphProperties(await repackDocx(cleaned, { updateModifiedDate: false })),
      ),
    ).toBe(before);
  });

  test("find-and-replace cloning preserves untouched paragraph properties", async () => {
    const source = await documentWithSourceProperties();
    const before = canonicalElement(await firstParagraphProperties(source));
    const parsed = await parseDocx(source, { preloadFonts: false });

    const replaced = replaceTextInDocument(
      parsed,
      { start: { paragraphIndex: 0, offset: 0 }, end: { paragraphIndex: 0, offset: 4 } },
      "Edited",
    );
    expect(firstParagraph(replaced).content).not.toEqual(firstParagraph(parsed).content);
    expect(getParagraphPropertySource(firstParagraph(replaced))).toEqual(
      getParagraphPropertySource(firstParagraph(parsed)),
    );
    expect(
      canonicalElement(
        await firstParagraphProperties(await repackDocx(replaced, { updateModifiedDate: false })),
      ),
    ).toBe(before);
  });

  test("public run consolidation preserves paragraph properties", async () => {
    const source = await documentWithSourceProperties();
    const before = canonicalElement(await firstParagraphProperties(source));
    const parsed = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const run = paragraph.content.find((item) => item.type === "run");
    if (!run) {
      panic("The fixture paragraph has no run.");
    }
    paragraph.content = [
      { ...run, content: [{ type: "text", text: "First" }] },
      { ...run, content: [{ type: "text", text: "Second" }] },
    ];

    const consolidated = consolidateParagraph(paragraph);
    expect(consolidated).not.toBe(paragraph);
    expect(getParagraphPropertySource(consolidated)).toEqual(getParagraphPropertySource(paragraph));
    parsed.package.document.content[0] = consolidated;
    expect(
      canonicalElement(
        await firstParagraphProperties(await repackDocx(parsed, { updateModifiedDate: false })),
      ),
    ).toBe(before);
  });

  test("load-time paraId allocation preserves an edited paragraph source that had no id", async () => {
    const source = await documentWithSourceProperties(SOURCE_PROPERTIES, null);
    const before = canonicalElement(await firstParagraphProperties(source));
    const parsed = await parseDocx(source, { preloadFonts: false });
    const proseDoc = ensureParaIdsInDoc(toProseDoc(parsed));
    const proseParagraph = proseDoc.child(0);
    const editedParagraph = proseParagraph.type.create(
      proseParagraph.attrs,
      proseParagraph.type.schema.text("Edited"),
    );
    const editedDoc = proseDoc.type.create(proseDoc.attrs, [editedParagraph]);
    const restored = fromProseDoc(editedDoc, parsed);

    expect(firstParagraph(restored).paraId).toMatch(/^[0-9A-F]{8}$/u);
    expect(getParagraphPropertySource(firstParagraph(restored))).toEqual(
      getParagraphPropertySource(firstParagraph(parsed)),
    );
    expect(
      canonicalElement(
        await firstParagraphProperties(await repackDocx(restored, { updateModifiedDate: false })),
      ),
    ).toBe(before);
  });

  test("exact private node identity distinguishes duplicate source ids", async () => {
    const parsed = await parseDocx(await documentWithDuplicateParagraphIds(), {
      preloadFonts: false,
    });
    const restored = fromProseDoc(toProseDoc(parsed), parsed);
    const paragraphs = restored.package.document.content.filter(
      (block): block is Paragraph => block.type === "paragraph",
    );
    expect(paragraphs).toHaveLength(2);
    expect(getParagraphPropertySource(paragraphs[0])?.xml).toContain('w:left="720"');
    expect(getParagraphPropertySource(paragraphs[1])?.xml).toContain('w:left="1440"');
  });

  test("a table crossing package ownership keeps its local properties without its durable token", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const sourceParagraph = firstParagraph(parsed);
    const proseDoc = toProseDoc(parsed);
    const sourceNode = proseDoc.child(0);
    const nodeSchema = proseDoc.type.schema;
    const template = nodeSchema.node("table", null, [
      nodeSchema.node("tableRow", null, [nodeSchema.node("tableCell", null, [sourceNode])]),
    ]);
    const copied = tableFromTemplate({ schema: nodeSchema, template });
    if (!copied) {
      panic("expected the table template to cross the package boundary");
    }
    const contractFree = nodeSchema.node("doc", null, [copied]);
    const restored = fromProseDoc(contractFree);
    const table = restored.package.document.content.at(0);
    const paragraph = table?.type === "table" ? table.rows.at(0)?.cells.at(0)?.content.at(0) : null;
    if (paragraph?.type !== "paragraph") {
      panic("expected the copied table paragraph");
    }

    expect(getParagraphPropertySource(paragraph)).toEqual(
      getParagraphPropertySource(sourceParagraph),
    );
    expect(getParagraphPropertySourceToken(paragraph)).toBeUndefined();
  });

  test("a hidden vertical-merge cell crosses with its local properties but no durable token", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const sourceParagraph = firstParagraph(parsed);
    const proseDoc = toProseDoc(parsed);
    const nodeSchema = proseDoc.type.schema;
    const hiddenCell = { type: "tableCell" as const, content: [sourceParagraph] };
    const template = nodeSchema.node("table", null, [
      nodeSchema.node("tableRow", null, [
        nodeSchema.node("tableCell", { rowspan: 2, _docxVMergeContinuationCells: [hiddenCell] }, [
          nodeSchema.node("paragraph"),
        ]),
      ]),
    ]);
    const copied = tableFromTemplate({ schema: nodeSchema, template });
    if (!copied) {
      panic("expected the vertical-merge table template to cross the package boundary");
    }
    const continuation = expectTableCellAttrs(
      copied.child(0).child(0),
    )._docxVMergeContinuationCells?.at(0);
    const paragraph = continuation?.content.at(0);
    if (paragraph?.type !== "paragraph") {
      panic("expected the hidden continuation paragraph");
    }

    expect(getParagraphPropertySource(paragraph)).toEqual(
      getParagraphPropertySource(sourceParagraph),
    );
    expect(getParagraphPropertySourceToken(paragraph)).toBeUndefined();
  });

  test("durable source tokens survive regenerated duplicate paragraph ids", async () => {
    const parsed = await parseDocx(await documentWithDuplicateParagraphIds(), {
      preloadFonts: false,
    });
    const normalized = ensureParaIdsInDoc(toProseDoc(parsed));
    const detached = normalized.type.create(
      normalized.attrs,
      normalized.content.content.map((paragraph) =>
        paragraph.type.create(paragraph.attrs, paragraph.content, paragraph.marks),
      ),
    );
    const restored = fromProseDoc(detached, parsed);
    const paragraphs = restored.package.document.content.filter(
      (block): block is Paragraph => block.type === "paragraph",
    );

    expect(paragraphs).toHaveLength(2);
    expect(getParagraphPropertySource(paragraphs[0])?.xml).toContain('w:left="720"');
    expect(getParagraphPropertySource(paragraphs[1])?.xml).toContain('w:left="1440"');
  });

  test("durable duplicate-id sources survive a shallow document derivation and PM reconstruction", async () => {
    const parsed = await parseDocx(await documentWithDuplicateParagraphIds(), {
      preloadFonts: false,
    });
    const derived = { ...parsed, package: { ...parsed.package } };
    const normalized = ensureParaIdsInDoc(toProseDoc(parsed));
    const reconstructed = normalized.type.schema.nodeFromJSON(normalized.toJSON());
    const restored = fromProseDoc(reconstructed, derived);
    const paragraphs = restored.package.document.content.filter(
      (block): block is Paragraph => block.type === "paragraph",
    );

    expect(getParagraphPropertySource(paragraphs[0])?.xml).toContain('w:left="720"');
    expect(getParagraphPropertySource(paragraphs[1])?.xml).toContain('w:left="1440"');
  });

  test("a durable id-less source survives a shallow document derivation and PM reconstruction", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(SOURCE_PROPERTIES, null), {
      preloadFonts: false,
    });
    const derived = { ...parsed, package: { ...parsed.package } };
    const normalized = ensureParaIdsInDoc(toProseDoc(parsed));
    const reconstructed = normalized.type.schema.nodeFromJSON(normalized.toJSON());
    const restored = fromProseDoc(reconstructed, derived);

    expect(getParagraphPropertySource(firstParagraph(restored))).toEqual(
      getParagraphPropertySource(firstParagraph(parsed)),
    );
  });

  test("paragraph id transfer cannot change durable source ownership", async () => {
    const parsed = await parseDocx(await documentWithDuplicateParagraphIds(), {
      preloadFonts: false,
    });
    const [first, second] = parsed.package.document.content.filter(
      (block): block is Paragraph => block.type === "paragraph",
    );
    if (!first || !second) {
      panic("fixture did not produce two paragraphs");
    }
    first.paraId = "11111111";
    second.paraId = "22222222";
    const proseDoc = toProseDoc(parsed);
    const firstProseParagraph = proseDoc.child(0);
    expect(firstProseParagraph.attrs["paraId"]).toBe("11111111");

    const detached = proseDoc.type.create(
      proseDoc.attrs,
      proseDoc.content.content.map((paragraph) =>
        paragraph.type.create(paragraph.attrs, paragraph.content, paragraph.marks),
      ),
    );
    const restored = fromProseDoc(detached, parsed);
    const paragraphs = restored.package.document.content.filter(
      (block): block is Paragraph => block.type === "paragraph",
    );
    expect(getParagraphPropertySource(paragraphs[0])?.xml).toContain('w:left="720"');
    expect(getParagraphPropertySource(paragraphs[1])?.xml).toContain('w:left="1440"');
  });

  test("duplicate editable paragraph identities cannot share a property capture", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const proseDoc = toProseDoc(parsed);
    const json = proseDoc.toJSON();
    const paragraph = json.content?.at(0);
    if (!paragraph) {
      panic("fixture did not produce an editable paragraph");
    }
    const duplicate = proseDoc.type.schema.nodeFromJSON({
      ...json,
      content: [paragraph, paragraph],
    });
    expect(() => fromProseDoc(duplicate, parsed)).toThrow(ParagraphPropertySourceValidationError);
  });

  test("one exact parser-linked paragraph node cannot lend its capture twice", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const proseDoc = toProseDoc(parsed);
    const paragraph = proseDoc.child(0);
    const duplicate = proseDoc.type.create(proseDoc.attrs, [paragraph, paragraph]);
    expect(() => fromProseDoc(duplicate, parsed)).toThrow(ParagraphPropertySourceValidationError);
  });

  test("a copied slice cannot borrow the exact owner's property capture", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const proseDoc = toProseDoc(parsed);
    const paragraph = proseDoc.child(0);
    const copied = paragraph.type.create(
      { ...paragraph.attrs, paraId: "87654321" },
      paragraph.content,
      paragraph.marks,
    );
    const copiedSlice = new Slice(Fragment.fromArray([paragraph, copied]), 0, 0);
    const duplicated = proseDoc.type.create(proseDoc.attrs, copiedSlice.content);
    expect(() => fromProseDoc(duplicated, parsed)).toThrow(ParagraphPropertySourceValidationError);
  });

  test("rejects a source-bearing model paired with a contract-free ProseMirror document", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const proseDoc = toProseDoc(parsed);
    const contractFree = proseDoc.type.create(
      { ...proseDoc.attrs, _docxParagraphSourceContract: null },
      proseDoc.content,
      proseDoc.marks,
    );

    expect(() => fromProseDoc(contractFree, parsed)).toThrow(
      ParagraphPropertySourceValidationError,
    );
  });

  test("rejects a contract-bearing ProseMirror document paired with a source-free model", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const original = toProseDoc(parsed);
    const proseDoc = original.type.schema.nodeFromJSON(original.toJSON());
    const sourceFree = structuredClone(parsed);

    expect(() => fromProseDoc(proseDoc, sourceFree)).toThrow(
      ParagraphPropertySourceValidationError,
    );
  });

  test("a JSON-cloned model falls back to its typed paragraph properties", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const clone = structuredClone(parsed);
    expect(getParagraphPropertySource(firstParagraph(clone))).toBeUndefined();

    const properties = canonicalElement(
      await firstParagraphProperties(await repackDocx(clone, { updateModifiedDate: false })),
    );
    expect(properties).toContain("<ind w:left=720>");
    expect(properties).toContain("<bCs >");
  });

  test("an inherited Strict namespace is converted before the property source is detached", async () => {
    const parsed = await parseDocx(await documentWithInheritedStrictProperties(), {
      preloadFonts: false,
    });
    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const properties = canonicalElement(await firstParagraphProperties(saved));

    expect(properties).toContain("<spacing w:before=240>");
    expect(properties).not.toContain("12pt");
  });

  test("capture strips revision names only in the main document namespace", async () => {
    const parsed = await parseDocx(
      await documentWithSourceProperties(
        '<w:pPr xmlns:x="urn:example:extension"><x:sectPr/><x:pPrChange/>' +
          '<x:rPr><w:ins w:id="1"/></x:rPr><w:rPr><x:ins/></w:rPr></w:pPr>',
      ),
      { preloadFonts: false },
    );
    const source = getParagraphPropertySource(firstParagraph(parsed));
    expect(source?.xml).toContain("<x:sectPr/>");
    expect(source?.xml).toContain("<x:pPrChange/>");
    expect(source?.xml).toContain('<x:rPr><w:ins w:id="1"/></x:rPr>');
    expect(source?.xml).toContain("<w:rPr><x:ins/></w:rPr>");
  });

  test("foreign extension children survive property replay", async () => {
    const parsed = await parseDocx(
      await documentWithSourceProperties(
        '<w:pPr xmlns:x="urn:example:extension"><w:ind w:left="720"/>' +
          "<x:extension><x:sectPr/><x:pPrChange/><x:rPr><x:ins/></x:rPr></x:extension>" +
          "</w:pPr>",
      ),
      { preloadFonts: false },
    );
    const { documentXml } = await unzipDocx(
      await repackDocx(parsed, { updateModifiedDate: false }),
    );
    expect(documentXml).toContain(
      "<x:extension><x:sectPr/><x:pPrChange/><x:rPr><x:ins/></x:rPr></x:extension>",
    );
  });

  test("a formatting edit invalidates the captured properties", async () => {
    const source = await documentWithSourceProperties();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    paragraph.formatting = { ...paragraph.formatting, alignment: "right" };

    const savedProperties = canonicalElement(
      await firstParagraphProperties(
        await repackDocx(parsed, {
          updateModifiedDate: false,
        }),
      ),
    );
    expect(savedProperties).toContain("<jc w:val=right>");
    expect(savedProperties).not.toContain("<cnfStyle ");
  });

  test("current structural and revision children are composed around the capture", async () => {
    const source = await documentWithSourceProperties();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    paragraph.pPrMark = {
      kind: "ins",
      info: { id: 7, author: "Reviewer" },
    };
    paragraph.sectionProperties = { pageWidth: 12_240, pageHeight: 15_840 };
    paragraph.propertyChanges = [
      {
        type: "paragraphPropertyChange",
        info: { id: 8, author: "Reviewer" },
        previousFormatting: { alignment: "left" },
      },
    ];

    const document = parseXmlDocument(
      `<w:document xmlns:w="${NAMESPACES.w}" xmlns:r="${NAMESPACES.r}" xmlns:w15="${NAMESPACES.w15}">${serializeParagraph(paragraph)}</w:document>`,
    );
    const savedParagraph = getChildElements(document).at(0);
    const properties = getChildElements(savedParagraph).find(
      (child) => getLocalName(child.name) === "pPr",
    );
    expect(getChildElements(properties).map((child) => getLocalName(child.name))).toEqual([
      "ind",
      "cnfStyle",
      "rPr",
      "sectPr",
      "pPrChange",
    ]);
    const runProperties = getChildElements(properties).find(
      (child) => getLocalName(child.name) === "rPr",
    );
    if (!runProperties) {
      panic("The saved paragraph has no run properties.");
    }
    expect(getChildElements(runProperties).map((child) => getLocalName(child.name))).toEqual([
      "ins",
      "bCs",
      "sz",
      "szCs",
      "noProof",
    ]);
  });

  test.each(["ins", "del", "moveFrom", "moveTo"] as const)(
    "the generated paragraph-mark particle keeps %s before authored run properties",
    async (kind) => {
      const parsed = await parseDocx(
        await documentWithSourceProperties(
          '<w:pPr><w:rPr><w:rFonts w:ascii="Arial"/><w:bCs/><w:sz w:val="21"/>' +
            "<w:noProof/></w:rPr></w:pPr>",
        ),
        { preloadFonts: false },
      );
      const paragraph = firstParagraph(parsed);
      paragraph.pPrMark = { kind, info: { id: 7, author: "Reviewer" } };

      const saved = await repackDocx(parsed, { updateModifiedDate: false });
      const properties = await firstParagraphProperties(saved);
      const runProperties = getChildElements(properties).find(
        (child) => getLocalName(child.name) === "rPr",
      );
      expect(getChildElements(runProperties).map((child) => getLocalName(child.name))).toEqual([
        kind,
        "rFonts",
        "bCs",
        "sz",
        "noProof",
      ]);
    },
  );

  test("a hostile dateUtc prefix binding cannot capture an appended paragraph mark", async () => {
    const utcDate = "2026-09-08T08:07:06Z";
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: propertySource.xml.replace("<w:pPr", '<w:pPr xmlns:w16du="urn:foreign"'),
    });
    paragraph.pPrMark = {
      kind: "ins",
      info: {
        id: 7,
        author: "Reviewer",
        utcDate: { attribute: "hostile:dateUtc", value: utcDate },
      },
    };

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);
    expect(documentXml).not.toContain('xmlns:w16du="urn:foreign"');
    expect(documentXml).toContain(`xmlns:w16du="${DATE_UTC_NAMESPACE_URI}"`);
    expect(documentXml).toContain(`w16du:dateUtc="${utcDate}"`);
    expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(firstParagraph(reopened).pPrMark?.info.utcDate?.value).toBe(utcDate);
  });

  test("a hostile dateUtc prefix binding cannot capture an appended property change", async () => {
    const utcDate = "2026-09-08T08:07:06Z";
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: propertySource.xml.replace("<w:pPr", '<w:pPr xmlns:w16du="urn:foreign"'),
    });
    paragraph.propertyChanges = [
      {
        type: "paragraphPropertyChange",
        info: {
          id: 8,
          author: "Reviewer",
          utcDate: { attribute: "hostile:dateUtc", value: utcDate },
        },
        previousFormatting: { alignment: "left" },
      },
    ];

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);
    expect(documentXml).not.toContain('xmlns:w16du="urn:foreign"');
    expect(documentXml).toContain(`xmlns:w16du="${DATE_UTC_NAMESPACE_URI}"`);
    expect(documentXml).toContain(`w16du:dateUtc="${utcDate}"`);
    expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(firstParagraph(reopened).propertyChanges?.at(0)?.info.utcDate?.value).toBe(utcDate);
  });

  test("an otherwise unused foreign dateUtc prefix binding remains replayable", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: propertySource.xml.replace("<w:pPr", '<w:pPr xmlns:w16du="urn:foreign"'),
    });

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);
    expect(documentXml).toContain('xmlns:w16du="urn:foreign"');
    expect(documentXml).toContain('<w:cnfStyle w:val="000000100000" w:oddHBand="1"/>');
    expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
  });

  test.each(["ins", "del", "moveFrom", "moveTo"] as const)(
    "an editable %s paragraph mark preserves its namespaced UTC timestamp",
    async (kind) => {
      const utcDate = "2026-09-08T08:07:06Z";
      const parsed = await parseDocx(
        await documentWithSourceProperties(
          `<w:pPr><w:rPr><w:${kind} xmlns:du="${DATE_UTC_NAMESPACE_URI}" ` +
            `w:id="7" w:author="Reviewer" du:dateUtc="${utcDate}"/></w:rPr></w:pPr>`,
        ),
        { preloadFonts: false },
      );
      const proseDoc = toProseDoc(parsed);
      const cloned = proseDoc.type.schema.nodeFromJSON(proseDoc.toJSON());
      const restored = fromProseDoc(cloned, parsed);

      expect(firstParagraph(restored).pPrMark?.info.utcDate).toEqual({
        attribute: "w16du:dateUtc",
        value: utcDate,
      });
      const { documentXml } = await unzipDocx(
        await repackDocx(restored, { updateModifiedDate: false }),
      );
      expect(documentXml).toContain(`w16du:dateUtc="${utcDate}"`);
      expect(documentXml).not.toMatch(/\sdu:dateUtc=/u);
    },
  );

  test("an injected sibling invalidates the captured properties", async () => {
    const source = await documentWithSourceProperties();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: `${propertySource.xml}<w:r><w:t>injected</w:t></w:r>`,
    });

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);
    expect(documentXml).not.toContain("injected");
  });

  test.each([
    ["XML declaration", '<?xml version="9.9"?>', "", 'version="9.9"'],
    ["processing instruction before the root", "<?folio replay?>", "", "folio replay"],
    ["processing instruction after the root", "", "<?folio replay?>", "folio replay"],
    ["comment", "<!-- replay -->", "", "<!-- replay -->"],
    ["document type", "<!DOCTYPE pPr>", "", "<!DOCTYPE pPr>"],
    ["extra root", "", "<w:r><w:t>injected root</w:t></w:r>", "injected root"],
    ["unclosed root", "", "", "replay-unclosed"],
    ["mismatched root", "", "", "<w:pPr></w:rPr>"],
    ["leading text", "replay-leading", "", "replay-leading"],
    ["trailing text", "", "replay-trailing", "replay-trailing"],
    ["trailing CDATA", "", "<![CDATA[replay-cdata]]>", "replay-cdata"],
    ["named entity", "", "", "custom"],
    ["invalid numeric entity", "", "", "&#0;"],
  ] as const)(
    "%s cannot escape the captured root or suppress current revisions",
    async (name, before, after, forbidden) => {
      const parsed = await parseDocx(await documentWithSourceProperties(), {
        preloadFonts: false,
      });
      const paragraph = firstParagraph(parsed);
      const propertySource = getParagraphPropertySource(paragraph);
      if (!propertySource) {
        panic("The parsed paragraph has no property source.");
      }
      assignParagraphPropertySource(paragraph, {
        ...propertySource,
        xml: adversarialParagraphPropertiesXml(name, propertySource.xml, before, after),
      });
      paragraph.pPrMark = { kind: "ins", info: { id: 96, author: "Reviewer" } };
      paragraph.sectionProperties = { pageWidth: 12_240, pageHeight: 15_840 };
      paragraph.propertyChanges = [
        {
          type: "paragraphPropertyChange",
          info: { id: 97, author: "Reviewer" },
          previousFormatting: { alignment: "left" },
        },
      ];

      const saved = await repackDocx(parsed, { updateModifiedDate: false });
      const { documentXml } = await unzipDocx(saved);
      const properties = await firstParagraphProperties(saved);
      const propertyNames = getChildElements(properties).map((child) => getLocalName(child.name));
      const runProperties = getChildElements(properties).find(
        (child) => getLocalName(child.name) === "rPr",
      );

      expect(documentXml).not.toContain(forbidden);
      expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
      expect(propertyNames).toContain("sectPr");
      expect(propertyNames).toContain("pPrChange");
      expect(getChildElements(runProperties).map((child) => getLocalName(child.name))).toContain(
        "ins",
      );
    },
  );

  test.each([
    ["section properties", '<w:sectPr w:rsidR="11111111"/>', "sectPr", "pPr"],
    [
      "paragraph property change",
      '<w:pPrChange w:id="91" w:author="Injected"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>',
      "pPrChange",
      "pPr",
    ],
    ["paragraph-mark insertion", '<w:ins w:id="92" w:author="Injected"/>', "ins", "rPr"],
    ["paragraph-mark deletion", '<w:del w:id="93" w:author="Injected"/>', "del", "rPr"],
  ] as const)(
    "an injected %s invalidates the captured properties",
    async (_name, injected, tag, parent) => {
      const source = await documentWithSourceProperties();
      const parsed = await parseDocx(source, { preloadFonts: false });
      const paragraph = firstParagraph(parsed);
      const propertySource = getParagraphPropertySource(paragraph);
      if (!propertySource) {
        panic("The parsed paragraph has no property source.");
      }
      assignParagraphPropertySource(paragraph, {
        ...propertySource,
        xml:
          parent === "rPr"
            ? propertySource.xml.replace("<w:rPr>", `<w:rPr>${injected}`)
            : propertySource.xml.replace("</w:pPr>", `${injected}</w:pPr>`),
      });

      const properties = await firstParagraphProperties(
        await repackDocx(parsed, { updateModifiedDate: false }),
      );
      const directNames = getChildElements(properties).map((child) => getLocalName(child.name));
      if (tag === "ins" || tag === "del") {
        const runPropertyNames = getChildElements(properties)
          .filter((child) => getLocalName(child.name) === "rPr")
          .flatMap((runProperties) =>
            getChildElements(runProperties).map((child) => getLocalName(child.name)),
          );
        expect(runPropertyNames).not.toContain(tag);
      } else {
        expect(directNames).not.toContain(tag);
      }
    },
  );

  test("a foreign root namespace invalidates the captured properties", async () => {
    const source = await documentWithSourceProperties();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: '<x:pPr xmlns:x="urn:example:not-wordprocessingml"><x:injected/></x:pPr>',
    });

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);
    expect(documentXml).not.toContain("injected");
  });

  test("a Strict root namespace is converted before replay", async () => {
    const source = await documentWithSourceProperties();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: '<x:pPr xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main"><x:keepNext/></x:pPr>',
    });

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);
    expect(documentXml).toContain(
      '<x:pPr xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:keepNext/></x:pPr>',
    );
    expect(documentXml).not.toContain("http://purl.oclc.org/ooxml/");
  });

  test("a foreign nested rPr cannot capture the current paragraph-mark change", async () => {
    const source = await documentWithSourceProperties();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: propertySource.xml.replace(
        "<w:rPr>",
        '<x:wrapper xmlns:x="urn:example:extension"><x:rPr/></x:wrapper><w:rPr>',
      ),
    });
    paragraph.pPrMark = { kind: "ins", info: { id: 94, author: "Reviewer" } };

    const properties = await firstParagraphProperties(
      await repackDocx(parsed, { updateModifiedDate: false }),
    );
    const runProperties = getChildElements(properties).find(
      (child) => getLocalName(child.name) === "rPr",
    );
    expect(getChildElements(runProperties).map((child) => getLocalName(child.name))).toContain(
      "ins",
    );
    expect(getChildElements(properties).map((child) => getLocalName(child.name))).toContain(
      "wrapper",
    );
  });

  test.each([
    [
      "section properties",
      '<mc:AlternateContent><mc:Choice Requires="w14"><w:sectPr w:rsidR="11111111"/></mc:Choice></mc:AlternateContent>',
      "sectPr",
    ],
    [
      "paragraph property change",
      '<mc:AlternateContent><mc:Choice Requires="w14"><w:pPrChange w:id="91" w:author="Injected"><w:pPr/></w:pPrChange></mc:Choice></mc:AlternateContent>',
      "pPrChange",
    ],
    [
      "paragraph-mark insertion",
      '<mc:AlternateContent><mc:Choice Requires="w14"><w:ins w:id="92" w:author="Injected"/></mc:Choice></mc:AlternateContent>',
      "ins",
    ],
  ] as const)("a nested injected %s invalidates the capture", async (_name, injected, tag) => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml:
        tag === "ins"
          ? propertySource.xml.replace("<w:rPr>", `<w:rPr>${injected}`)
          : propertySource.xml.replace("</w:pPr>", `${injected}</w:pPr>`),
    });
    paragraph.pPrMark = { kind: "ins", info: { id: 98, author: "Current" } };

    const properties = await firstParagraphProperties(
      await repackDocx(parsed, { updateModifiedDate: false }),
    );
    const serialized = JSON.stringify(properties);
    expect(serialized).not.toContain("Injected");
    expect(serialized).toContain("Current");
  });

  test.each([
    "pPrChange",
    "sectPr",
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "cellDel",
    "cellIns",
    "cellMerge",
    "numberingChange",
    "tblGridChange",
    "tblPrChange",
    "tcPrChange",
    "trPrChange",
    "rPrChange",
  ] as const)(
    "a nested WordprocessingML %s cannot bypass the modeled composition slots",
    async (revisionName) => {
      const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
      const paragraph = firstParagraph(parsed);
      const propertySource = getParagraphPropertySource(paragraph);
      if (!propertySource) {
        panic("The parsed paragraph has no property source.");
      }
      assignParagraphPropertySource(paragraph, {
        ...propertySource,
        xml: propertySource.xml.replace(
          '<w:ind w:left="720" w:leftChars="100"/>',
          `<w:ind w:left="720" w:leftChars="100"><w:${revisionName} ` +
            'w:author="Poison"/></w:ind>',
        ),
      });
      paragraph.pPrMark = { kind: "moveTo", info: { id: 98, author: "Current" } };
      paragraph.sectionProperties = { pageWidth: 12_240, pageHeight: 15_840 };
      paragraph.propertyChanges = [
        {
          type: "paragraphPropertyChange",
          info: { id: 99, author: "Current" },
          previousFormatting: { alignment: "left" },
        },
      ];

      const saved = await repackDocx(parsed, { updateModifiedDate: false });
      const { documentXml } = await unzipDocx(saved);
      expect(documentXml).not.toContain("Poison");
      expect(documentXml).toContain('w:author="Current"');
      expect(documentXml).toContain("<w:moveTo ");
      expect(documentXml).toContain("<w:sectPr");
      expect(documentXml).toContain("<w:pPrChange ");
      expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
    },
  );

  test.each(["p", "tbl", "r", "t"] as const)(
    "a nested WordprocessingML %s cannot hide beneath a paragraph property",
    async (contentName) => {
      const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
      const paragraph = firstParagraph(parsed);
      const propertySource = getParagraphPropertySource(paragraph);
      if (!propertySource) {
        panic("The parsed paragraph has no property source.");
      }
      assignParagraphPropertySource(paragraph, {
        ...propertySource,
        xml: propertySource.xml.replace(
          '<w:ind w:left="720" w:leftChars="100"/>',
          `<w:ind w:left="720" w:leftChars="100"><w:${contentName}>POISON</w:${contentName}></w:ind>`,
        ),
      });

      const saved = await repackDocx(parsed, { updateModifiedDate: false });
      const { documentXml } = await unzipDocx(saved);
      expect(documentXml).not.toContain("POISON");
      expect(documentXml).toContain('w:left="720"');
      expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
    },
  );

  test("non-whitespace text cannot hide beneath a paragraph property", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: propertySource.xml.replace(
        '<w:ind w:left="720" w:leftChars="100"/>',
        '<w:ind w:left="720" w:leftChars="100">POISON</w:ind>',
      ),
    });

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);
    expect(documentXml).not.toContain("POISON");
    expect(documentXml).toContain('w:left="720"');
    expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
  });

  test.each([
    ["canonical prefix", '<w:ind w:left="9999"/>'],
    ["alternate prefix", `<x:ind xmlns:x="${NAMESPACES.w}" x:left="9999"/>`],
  ] as const)(
    "modeled properties under an MC branch cannot shadow the current model (%s)",
    async (_name, injected) => {
      const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
      const paragraph = firstParagraph(parsed);
      const propertySource = getParagraphPropertySource(paragraph);
      if (!propertySource) {
        panic("The parsed paragraph has no property source.");
      }
      assignParagraphPropertySource(paragraph, {
        ...propertySource,
        xml: propertySource.xml.replace(
          "</w:pPr>",
          `<mc:AlternateContent><mc:Choice Requires="w14">${injected}</mc:Choice>` +
            `<mc:Fallback><w:ind w:left="8888"/></mc:Fallback></mc:AlternateContent></w:pPr>`,
        ),
      });
      paragraph.pPrMark = { kind: "ins", info: { id: 101, author: "Current" } };

      const saved = await repackDocx(parsed, { updateModifiedDate: false });
      const { documentXml } = await unzipDocx(saved);
      expect(documentXml).not.toContain('w:left="9999"');
      expect(documentXml).not.toContain('w:left="8888"');
      expect(documentXml).toContain('w:left="720"');
      expect(documentXml).toContain('w:author="Current"');
      expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
    },
  );

  test("an extension-only MC branch remains a fixed point", async () => {
    const extension =
      '<mc:AlternateContent xmlns:x="urn:example:paragraph-extension">' +
      '<mc:Choice Requires="x"><x:property x:value="kept"/></mc:Choice>' +
      "<mc:Fallback/></mc:AlternateContent>";
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: propertySource.xml.replace("</w:pPr>", `${extension}</w:pPr>`),
    });

    const first = await repackDocx(parsed, { updateModifiedDate: false });
    const reparsed = await parseDocx(first, { preloadFonts: false });
    const second = await repackDocx(reparsed, { updateModifiedDate: false });
    const firstProperties = canonicalElement(await firstParagraphProperties(first));
    const secondProperties = canonicalElement(await firstParagraphProperties(second));

    expect(firstProperties).toContain("property");
    expect(secondProperties).toBe(firstProperties);
  });

  test("direct block content cannot replay inside paragraph properties", async () => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: propertySource.xml.replace(
        "</w:pPr>",
        "<w:p><w:r><w:t>POISON</w:t></w:r></w:p></w:pPr>",
      ),
    });

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);
    expect(documentXml).not.toContain("POISON");
    expect(documentXml).toContain("Text");
    expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
  });

  test("an unmodeled paragraph-mark rPrChange is ineligible for captured full-repack replay", async () => {
    const parsed = await parseDocx(
      await documentWithSourceProperties(
        '<w:pPr><w:rPr><w:bCs/><w:sz w:val="21"/><w:noProof/>' +
          '<w:rPrChange w:id="12" w:author="Original"><w:rPr><w:b/></w:rPr>' +
          "</w:rPrChange></w:rPr></w:pPr>",
      ),
      { preloadFonts: false },
    );
    const propertySource = getParagraphPropertySource(firstParagraph(parsed));
    expect(propertySource?.xml).toContain("<w:rPrChange ");
    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const { documentXml } = await unzipDocx(saved);

    expect(documentXml).not.toContain("<w:rPrChange ");
    expect(documentXml).not.toContain('w:author="Original"');
    expect(documentXml).toContain("<w:bCs/>");
    expect(documentXml).toContain('<w:sz w:val="21"/>');
    expect(isSafeCapturedXmlDocument(documentXml)).toBe(true);
  });

  test("a wide extension property tree is validated without argument spreading", async () => {
    const leaves = Array.from({ length: 70_000 }, () => "<x:leaf/>").join("");
    const parsed = await parseDocx(
      await documentWithSourceProperties(
        `<w:pPr><w:ind w:left="720"/><x:extension xmlns:x="urn:wide">${leaves}</x:extension></w:pPr>`,
      ),
      { preloadFonts: false },
    );

    const xml = serializeParagraph(firstParagraph(parsed));
    expect(xml.match(/<x:leaf\/>/gu)?.length).toBe(70_000);
    expect(
      isSafeCapturedXmlDocument(
        `<w:document xmlns:w="${NAMESPACES.w}" xmlns:w14="${NAMESPACES.w14}">${xml}</w:document>`,
      ),
    ).toBe(true);
  });

  test.each([
    ["r", NAMESPACES.r],
    ["w15", NAMESPACES.w15],
  ] as const)("a shadowed %s prefix cannot capture appended section markup", async (prefix) => {
    const parsed = await parseDocx(await documentWithSourceProperties(), { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, {
      ...propertySource,
      xml: propertySource.xml.replace("<w:pPr>", `<w:pPr xmlns:${prefix}="urn:foreign">`),
    });
    paragraph.sectionProperties = {
      headerReferences: [{ type: "default", rId: "rId1" }],
      footnoteColumns: 2,
    };

    const document = parseXmlDocument(
      `<w:document xmlns:w="${NAMESPACES.w}" xmlns:r="${NAMESPACES.r}" xmlns:w15="${NAMESPACES.w15}">${serializeParagraph(paragraph)}</w:document>`,
    );
    const savedParagraph = getChildElements(document).at(0);
    const properties = getChildElements(savedParagraph).find(
      (child) => getLocalName(child.name) === "pPr",
    );
    const section = getChildElements(properties).find(
      (child) => getLocalName(child.name) === "sectPr",
    );
    const headerReference = getChildElements(section).find(
      (child) => getLocalName(child.name) === "headerReference",
    );
    const footnoteColumns = getChildElements(section).find(
      (child) => getLocalName(child.name) === "footnoteColumns",
    );
    expect(getAttributeByNamespaceUri(headerReference, new Set([NAMESPACES.r]), "id")).toBe("rId1");
    expect(footnoteColumns?.namespaceUri).toBe(NAMESPACES.w15);
  });

  test("fallback paragraph properties retain schema child order", () => {
    const xml = serializeParagraphFormatting({
      suppressLineNumbers: true,
      borders: { top: { style: "single", size: 4, color: "000000" } },
      shading: { fill: "FFFFFF" },
      tabs: [{ type: "left", position: 720 }],
    });
    expect(xml.indexOf("<w:suppressLineNumbers")).toBeLessThan(xml.indexOf("<w:pBdr>"));
    expect(xml.indexOf("<w:pBdr>")).toBeLessThan(xml.indexOf("<w:shd"));
    expect(xml.indexOf("<w:shd")).toBeLessThan(xml.indexOf("<w:tabs>"));
  });

  test.each([
    [
      "property root",
      `<x:pPr xmlns:x="${NAMESPACES.w}" xmlns:w="urn:example:foreign"><x:rPr><x:bCs/></x:rPr></x:pPr>`,
    ],
    [
      "paragraph-mark properties",
      `<w:pPr><x:rPr xmlns:x="${NAMESPACES.w}" xmlns:w="urn:example:foreign"><x:bCs/></x:rPr></w:pPr>`,
    ],
  ] as const)("the %s cannot rebind the serializer's w prefix", async (_name, xml) => {
    const source = await documentWithSourceProperties();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraph(parsed);
    const propertySource = getParagraphPropertySource(paragraph);
    if (!propertySource) {
      panic("The parsed paragraph has no property source.");
    }
    assignParagraphPropertySource(paragraph, { ...propertySource, xml });
    paragraph.pPrMark = { kind: "ins", info: { id: 95, author: "Reviewer" } };

    const properties = await firstParagraphProperties(
      await repackDocx(parsed, { updateModifiedDate: false }),
    );
    const runProperties = getChildElements(properties).find(
      (child) => getLocalName(child.name) === "rPr",
    );
    const insertion = getChildElements(runProperties).find(
      (child) => getLocalName(child.name) === "ins",
    );
    expect(properties.namespaceUri).toBe(NAMESPACES.w);
    expect(runProperties?.namespaceUri).toBe(NAMESPACES.w);
    expect(insertion?.namespaceUri).toBe(NAMESPACES.w);
  });
});
