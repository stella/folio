import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import JSZip from "jszip";

import { modelParagraphFormattingEmission } from "../internal/paragraphFormattingSerialization";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, ParagraphFormatting } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import { parseDocx } from "./parser";
import {
  assignParagraphPropertySource,
  copyParagraphPropertyCapture,
  getParagraphPropertySource,
  paragraphPropertySourceMatchesEmission,
} from "./paragraphPropertySource";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { serializeParagraph, serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { attemptSelectiveSave } from "./selectiveSave";

const PARAGRAPH_ID = "12345678";
const NUM_PR = { kind: "reference", numId: 7, ilvl: 1 } as const;
const OTHER_NUM_PR = { kind: "reference", numId: 8, ilvl: 1 } as const;
const UNKNOWN_PROPERTY = '<x:unknown x:value="preserve-me"/>';
const NUMBERING_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

type SourceNumbering = "absent" | "direct" | "style-sourced";

type ProvenanceCase = Readonly<{
  name: string;
  sourceNumbering: SourceNumbering;
  mutate: (formatting: ParagraphFormatting) => void;
  preservesCapture: boolean;
  emitsDirectNumbering: boolean;
  expectedFragments?: readonly string[];
}>;

const CASES = [
  {
    name: "absent numbering plus unused style provenance",
    sourceNumbering: "absent",
    mutate: (formatting) => {
      formatting.numPrFromStyle = NUM_PR;
    },
    preservesCapture: true,
    emitsDirectNumbering: false,
  },
  {
    name: "emission-empty dependent and compound formatting",
    sourceNumbering: "absent",
    mutate: (formatting) => {
      formatting.hangingIndent = true;
      formatting.borders = {};
      formatting.tabs = [];
      formatting.frame = {};
      formatting.runProperties = {
        fontFamily: {},
        styleId: "",
        language: {},
        color: {},
      };
      // `runInWithNext = false` used to belong here: it emitted nothing,
      // because the serializer wrote `w:specVanish` only for an on. It now
      // emits the explicit off, so it is a modelled value like any other and
      // the capture it disagrees with is stale.
    },
    preservesCapture: true,
    emitsDirectNumbering: false,
  },
  {
    name: "direct numbering plus unequal style provenance",
    sourceNumbering: "direct",
    mutate: (formatting) => {
      formatting.numPrFromStyle = OTHER_NUM_PR;
    },
    preservesCapture: true,
    emitsDirectNumbering: true,
  },
  {
    name: "direct numbering plus equal style provenance",
    sourceNumbering: "direct",
    mutate: (formatting) => {
      formatting.numPrFromStyle = NUM_PR;
    },
    preservesCapture: false,
    emitsDirectNumbering: false,
  },
  {
    name: "style-sourced numbering plus spacing provenance",
    sourceNumbering: "style-sourced",
    mutate: (formatting) => {
      formatting.spacingExplicit = { before: true };
    },
    preservesCapture: true,
    emitsDirectNumbering: false,
  },
  {
    name: "style-sourced numbering made direct",
    sourceNumbering: "style-sourced",
    mutate: (formatting) => {
      delete formatting.numPrFromStyle;
    },
    preservesCapture: false,
    emitsDirectNumbering: true,
  },
  {
    name: "nested paragraph border formatting",
    sourceNumbering: "absent",
    mutate: (formatting) => {
      formatting.borders = {
        bottom: {
          style: "single",
          size: 8,
          space: 1,
          color: { themeColor: "accent1", themeTint: "80" },
        },
      };
    },
    preservesCapture: false,
    emitsDirectNumbering: false,
    expectedFragments: [
      '<w:bottom w:val="single" w:sz="8" w:space="1" w:themeColor="accent1" w:themeTint="80"/>',
    ],
  },
  {
    name: "nested paragraph shading and tab formatting",
    sourceNumbering: "absent",
    mutate: (formatting) => {
      formatting.shading = {
        pattern: "clear",
        color: { rgb: "202020" },
        fill: { themeColor: "accent2", themeShade: "40" },
      };
      formatting.tabs = [{ position: 720, alignment: "right", leader: "dot" }];
    },
    preservesCapture: false,
    emitsDirectNumbering: false,
    expectedFragments: [
      '<w:shd w:val="clear" w:color="202020" w:themeFill="accent2" w:themeFillShade="40"/>',
      '<w:tab w:val="right" w:pos="720" w:leader="dot"/>',
    ],
  },
  {
    name: "nested frame formatting",
    sourceNumbering: "absent",
    mutate: (formatting) => {
      formatting.frame = {
        width: 1_440,
        height: 720,
        x: 0,
        y: 0,
        hAnchor: "margin",
        vAnchor: "page",
        wrap: "around",
      };
    },
    preservesCapture: false,
    emitsDirectNumbering: false,
    expectedFragments: [
      '<w:framePr w:w="1440" w:h="720" w:hAnchor="margin" w:vAnchor="page" w:x="0" w:y="0" w:wrap="around"/>',
    ],
  },
  {
    name: "nested paragraph-mark run formatting",
    sourceNumbering: "absent",
    mutate: (formatting) => {
      formatting.runProperties = {
        fontFamily: { ascii: "Arial", hint: "default", asciiTheme: "minorAscii" },
        color: { themeColor: "accent3", themeTint: "20" },
        underline: {
          style: "single",
          color: { themeColor: "accent4", themeShade: "40" },
        },
        shading: { pattern: "clear", fill: { rgb: "E0E0E0" } },
        language: { val: "en-US", eastAsia: "ja-JP", bidi: "ar-SA" },
      };
    },
    preservesCapture: false,
    emitsDirectNumbering: false,
    expectedFragments: [
      '<w:rFonts w:ascii="Arial" w:hint="default" w:asciiTheme="minorAscii"/>',
      '<w:color w:themeColor="accent3" w:themeTint="20"/>',
      '<w:u w:val="single" w:themeColor="accent4" w:themeShade="40"/>',
      '<w:shd w:val="clear" w:fill="E0E0E0"/>',
      '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>',
    ],
  },
] as const satisfies readonly ProvenanceCase[];

const paragraphProperties = (numbering: SourceNumbering): string => {
  const numberingXml =
    numbering === "direct" ? '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr>' : "";
  const styleXml = numbering === "style-sourced" ? '<w:pStyle w:val="Numbered"/>' : "";
  return `<w:pPr xmlns:x="urn:folio:unknown-property">${styleXml}${numberingXml}${UNKNOWN_PROPERTY}</w:pPr>`;
};

const generatedDocument = async (numbering: SourceNumbering): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
  const documentRelationships = await zip.file("word/_rels/document.xml.rels")?.async("text");
  const documentXml = await zip.file("word/document.xml")?.async("text");
  const stylesXml = await zip.file("word/styles.xml")?.async("text");
  if (!contentTypes || !documentRelationships || !documentXml || !stylesXml) {
    panic("The generated package is missing a required part.");
  }

  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(
      "</Types>",
      `  <Override PartName="/word/numbering.xml" ContentType="${NUMBERING_CONTENT_TYPE}"/>\n</Types>`,
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    documentRelationships.replace(
      "</Relationships>",
      `  <Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.numbering}" Target="numbering.xml"/>\n</Relationships>`,
    ),
  );
  zip.file(
    "word/styles.xml",
    stylesXml.replace(
      "</w:styles>",
      '<w:style w:type="paragraph" w:styleId="Numbered"><w:name w:val="Numbered"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr></w:style></w:styles>',
    ),
  );
  zip.file("word/numbering.xml", NUMBERING_XML);
  zip.file(
    "word/document.xml",
    documentXml
      .replace(
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
      )
      .replace(
        /<w:body>[\s\S]*<\/w:body>/u,
        `<w:body><w:p w14:paraId="${PARAGRAPH_ID}">${paragraphProperties(numbering)}<w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body>`,
      ),
  );

  return zip.generateAsync({ type: "arraybuffer" });
};

const firstParagraph = (document: Document): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    panic("The generated document has no first paragraph.");
  }
  return paragraph;
};

const firstParagraphPropertiesXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (!documentXml) {
    panic("The saved package has no main document part.");
  }
  return /<w:pPr(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/w:pPr>)/u.exec(documentXml)?.at(0) ?? "";
};

const mutateParsedParagraph = (document: Document, fixture: ProvenanceCase): Paragraph => {
  const paragraph = firstParagraph(document);
  paragraph.formatting ??= {};
  fixture.mutate(paragraph.formatting);
  return paragraph;
};

describe("captured paragraph properties follow modeled fallback emission", () => {
  test("owned captures replay exactly while their modeled emission is current", () => {
    const formattingArbitrary = fc.record({
      alignment: fc.constantFrom("left", "center", "right", "both"),
      keepNext: fc.boolean(),
      numPr: fc.option(
        fc.record({
          kind: fc.constant("reference" as const),
          numId: fc.integer({ min: 1, max: 100 }),
          ilvl: fc.integer({ min: 0, max: 8 }),
        }),
        { nil: undefined },
      ),
    });

    fc.assert(
      fc.property(formattingArbitrary, fc.boolean(), (baseFormatting, keepsEmission) => {
        const paragraph: Paragraph = {
          type: "paragraph",
          formatting: baseFormatting,
          content: [{ type: "run", content: [{ type: "text", text: "Text" }] }],
        };
        const fallbackProperties = serializeParagraphFormatting(baseFormatting);
        const sourceProperties = fallbackProperties.replace(
          "<w:pPr>",
          `<w:pPr xmlns:x="urn:folio:unknown-property">${UNKNOWN_PROPERTY}`,
        );
        assignParagraphPropertySource(paragraph, sourceProperties);
        const source = getParagraphPropertySource(paragraph);
        if (!source) {
          panic("The generated paragraph has no owned property source.");
        }

        paragraph.formatting = keepsEmission
          ? { ...baseFormatting, spacingExplicit: { before: true } }
          : { ...baseFormatting, keepNext: !baseFormatting.keepNext };
        const currentEmission = modelParagraphFormattingEmission(paragraph.formatting);

        expect(Object.isFrozen(source)).toBe(true);
        expect(paragraphPropertySourceMatchesEmission(source, currentEmission)).toBe(keepsEmission);
        const serialized = serializeParagraph(paragraph);
        expect(serialized.includes(UNKNOWN_PROPERTY)).toBe(keepsEmission);
        if (!keepsEmission) {
          expect(serialized).toContain(serializeParagraphFormatting(paragraph.formatting));
        }
      }),
      { numRuns: 128 },
    );
  });

  test("copying a capture onto an edited paragraph preserves the source emission identity", () => {
    const source: Paragraph = {
      type: "paragraph",
      formatting: { keepNext: false },
      content: [{ type: "run", content: [{ type: "text", text: "Text" }] }],
    };
    assignParagraphPropertySource(
      source,
      '<w:pPr><w:keepNext w:val="0"/><x:stale xmlns:x="urn:folio:test"/></w:pPr>',
    );
    const target: Paragraph = {
      ...source,
      formatting: { keepNext: true },
    };

    copyParagraphPropertyCapture(target, source);

    const copied = getParagraphPropertySource(target);
    if (!copied) {
      panic("The copied paragraph has no captured property source.");
    }
    expect(Object.isFrozen(copied)).toBe(true);
    expect(
      paragraphPropertySourceMatchesEmission(
        copied,
        modelParagraphFormattingEmission(target.formatting),
      ),
    ).toBe(false);
    const serialized = serializeParagraph(target);
    expect(serialized).toContain("<w:keepNext/>");
    expect(serialized).not.toContain('w:keepNext w:val="0"');
    expect(serialized).not.toContain("x:stale");
  });

  test.each(CASES)("$name", async (fixture) => {
    const sourceBuffer = await generatedDocument(fixture.sourceNumbering);
    const fullDocument = await parseDocx(sourceBuffer, { preloadFonts: false });
    const selectiveDocument = await parseDocx(sourceBuffer, { preloadFonts: false });
    const fullParagraph = firstParagraph(fullDocument);
    const selectiveParagraph = firstParagraph(selectiveDocument);
    const source = getParagraphPropertySource(fullParagraph);
    const selectiveSource = getParagraphPropertySource(selectiveParagraph);
    if (!source || !selectiveSource) {
      panic("The generated paragraph has no captured property source.");
    }
    expect(source.xml).toContain(UNKNOWN_PROPERTY);
    expect(selectiveSource).toEqual(source);
    expect(Object.isFrozen(source)).toBe(true);
    expect(Object.isFrozen(selectiveSource)).toBe(true);

    const beforeFormattingJson = canonicalJson(fullParagraph.formatting ?? {});
    const beforeEmission = modelParagraphFormattingEmission(fullParagraph.formatting);
    expect(paragraphPropertySourceMatchesEmission(source, beforeEmission)).toBe(true);

    mutateParsedParagraph(fullDocument, fixture);
    mutateParsedParagraph(selectiveDocument, fixture);
    const afterFormattingJson = canonicalJson(fullParagraph.formatting ?? {});
    const afterEmission = modelParagraphFormattingEmission(fullParagraph.formatting);
    expect(afterFormattingJson).not.toBe(beforeFormattingJson);
    expect(modelParagraphFormattingEmission(selectiveParagraph.formatting)).toEqual(afterEmission);
    if (fixture.preservesCapture) {
      expect(afterEmission).toEqual(beforeEmission);
    } else {
      expect(afterEmission).not.toEqual(beforeEmission);
    }
    expect(paragraphPropertySourceMatchesEmission(source, afterEmission)).toBe(
      fixture.preservesCapture,
    );

    const [full, selective] = await Promise.all([
      repackDocx(fullDocument, { updateModifiedDate: false }),
      attemptSelectiveSave(selectiveDocument, sourceBuffer, {
        changedParaIds: new Set([PARAGRAPH_ID]),
        structuralChange: false,
        hasUntrackedChanges: false,
      }),
    ]);
    expect(selective).not.toBeNull();
    if (!selective) {
      panic("Selective save rejected the generated paragraph mutation.");
    }

    const [fullProperties, selectiveProperties] = await Promise.all([
      firstParagraphPropertiesXml(full),
      firstParagraphPropertiesXml(selective),
    ]);
    expect(selectiveProperties).toBe(fullProperties);
    expect(fullProperties.includes(UNKNOWN_PROPERTY)).toBe(fixture.preservesCapture);
    expect(fullProperties.includes("<w:numPr")).toBe(fixture.emitsDirectNumbering);
    for (const fragment of fixture.expectedFragments ?? []) {
      expect(fullProperties).toContain(fragment);
    }

    const [parsedFull, parsedSelective] = await Promise.all([
      parseDocx(full, { preloadFonts: false }),
      parseDocx(selective, { preloadFonts: false }),
    ]);
    const fullEmission = modelParagraphFormattingEmission(firstParagraph(parsedFull).formatting);
    const selectiveEmission = modelParagraphFormattingEmission(
      firstParagraph(parsedSelective).formatting,
    );
    expect(fullEmission).toEqual(afterEmission);
    expect(selectiveEmission).toEqual(afterEmission);

    const fixedPointBuffer = await repackDocx(parsedFull, { updateModifiedDate: false });
    const fixedPointProperties = await firstParagraphPropertiesXml(fixedPointBuffer);
    const fixedPointDocument = await parseDocx(fixedPointBuffer, { preloadFonts: false });
    expect(fixedPointProperties).toBe(fullProperties);
    expect(modelParagraphFormattingEmission(firstParagraph(fixedPointDocument).formatting)).toEqual(
      fullEmission,
    );
  });

  test("editable conversion restores style-numbering provenance through the shared model", async () => {
    const parsed = await parseDocx(await generatedDocument("style-sourced"), {
      preloadFonts: false,
    });
    const sourceParagraph = firstParagraph(parsed);
    const source = getParagraphPropertySource(sourceParagraph);
    if (!source) {
      panic("The generated style-numbered paragraph has no captured property source.");
    }

    const restored = firstParagraph(fromProseDoc(toProseDoc(parsed), parsed));

    expect(sourceParagraph.formatting?.numPr).toEqual(NUM_PR);
    expect(sourceParagraph.formatting?.numPrFromStyle).toEqual(NUM_PR);
    expect(restored.formatting?.numPr).toEqual(NUM_PR);
    expect(restored.formatting?.numPrFromStyle).toEqual(NUM_PR);
    expect(modelParagraphFormattingEmission(restored.formatting)).toEqual(
      modelParagraphFormattingEmission(sourceParagraph.formatting),
    );
    expect(getParagraphPropertySource(restored)).toEqual(source);
  });
});
