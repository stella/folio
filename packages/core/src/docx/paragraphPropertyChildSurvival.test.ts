/**
 * A paragraph property set keeps every child its content model declares.
 *
 * `CT_PPrBase` declares thirty-three properties and folio models fourteen of
 * them. The rest reached disk only while the whole `w:pPr` was replayed as
 * bytes, so the first edit to any paragraph property — an alignment command,
 * a spacing change, a style applied — rebuilt the element without them. The
 * sample table below is total over the declared set, so a child folio starts
 * declaring cannot join without an example that has to survive.
 *
 * The assertions are about a **capture-free** save: the paragraph's formatting
 * is edited before the save, which is exactly what invalidates the replay and
 * forces the writer to rebuild the element from the model.
 */

import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";
import { CONTAINER_CHILDREN, type DeclaredChild } from "./containerChildren.gen";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { unzipDocx } from "./unzip";

/**
 * One example per child `CT_PPr` declares, in the order it declares them.
 *
 * `Record`, never `Partial<Record>`: the table mirrors the generated
 * declared-child set, and a mirror that may be incomplete is a mirror that
 * drifts. Three of the entries — `w:rPr`, `w:sectPr` and `w:pPrChange` — are
 * the children the reader marks `OWNED_ELSEWHERE`, and they are here for the
 * same reason: their owners have to write them back too.
 */
const DECLARED_CHILD_SAMPLES = {
  pStyle: '<w:pStyle w:val="Heading1"/>',
  keepNext: "<w:keepNext/>",
  keepLines: "<w:keepLines/>",
  pageBreakBefore: '<w:pageBreakBefore w:val="0"/>',
  // Attributes in the order the writer emits them: a modelled child comes
  // back rebuilt from the model, so the sample is what the model spells.
  framePr: '<w:framePr w:dropCap="drop" w:lines="3" w:vAnchor="text" w:wrap="around"/>',
  widowControl: "<w:widowControl/>",
  numPr: '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="4"/></w:numPr>',
  suppressLineNumbers: "<w:suppressLineNumbers/>",
  pBdr: '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>',
  shd: '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>',
  tabs: '<w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs>',
  suppressAutoHyphens: "<w:suppressAutoHyphens/>",
  kinsoku: "<w:kinsoku/>",
  wordWrap: '<w:wordWrap w:val="0"/>',
  overflowPunct: "<w:overflowPunct/>",
  topLinePunct: "<w:topLinePunct/>",
  autoSpaceDE: '<w:autoSpaceDE w:val="0"/>',
  autoSpaceDN: '<w:autoSpaceDN w:val="0"/>',
  bidi: "<w:bidi/>",
  adjustRightInd: '<w:adjustRightInd w:val="0"/>',
  snapToGrid: "<w:snapToGrid/>",
  spacing: '<w:spacing w:before="240" w:after="120"/>',
  ind: '<w:ind w:left="720" w:firstLine="360"/>',
  contextualSpacing: "<w:contextualSpacing/>",
  mirrorIndents: "<w:mirrorIndents/>",
  suppressOverlap: "<w:suppressOverlap/>",
  jc: '<w:jc w:val="center"/>',
  textDirection: '<w:textDirection w:val="tbRl"/>',
  textAlignment: '<w:textAlignment w:val="baseline"/>',
  textboxTightWrap: '<w:textboxTightWrap w:val="allLines"/>',
  outlineLvl: '<w:outlineLvl w:val="2"/>',
  divId: '<w:divId w:val="1234567"/>',
  cnfStyle: '<w:cnfStyle w:val="000000100000" w:oddHBand="1"/>',
  rPr: '<w:rPr><w:sz w:val="21"/></w:rPr>',
  sectPr: '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
  pPrChange:
    '<w:pPrChange w:id="9" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:pPr><w:jc w:val="both"/></w:pPr></w:pPrChange>',
} as const satisfies Record<DeclaredChild<"paragraph-properties">, string>;

/** The markup a child must still be spelled with after the round trip. */
const evidenceOf = (child: DeclaredChild<"paragraph-properties">): string => {
  // A container with children comes back with the same opening tag; asserting
  // on the whole element would also assert that nothing inside it moved, which
  // is a different law and a different fixture.
  const sample = DECLARED_CHILD_SAMPLES[child];
  const close = sample.indexOf(">");
  return sample.endsWith("/>") ? sample : sample.slice(0, close + 1);
};

const CHILD_NAMES = CONTAINER_CHILDREN["paragraph-properties"];

/**
 * Each loop below builds, parses and saves one package per declared child, so
 * the default five seconds is a coin toss on a loaded machine rather than a
 * statement about the code.
 */
const ONE_PACKAGE_PER_DECLARED_CHILD_MS = 60_000;

/**
 * Children a `CT_PPrGeneral` owner does not reach.
 *
 * `w:rPr` and `w:sectPr` are not declared by `CT_PPrGeneral` at all.
 * `w:pPrChange` is, and folio has no owner for it there: the reader names the
 * paragraph's property-change parser as its owner, and the style tier has no
 * equivalent. Nothing is lost today, because a repack copies
 * `word/styles.xml` through byte for byte, but a rebuilt styles part would
 * drop it — recorded here rather than left for the rebuild to discover.
 */
const CHILDREN_OUTSIDE_THE_GENERAL_SET: ReadonlySet<string> = new Set([
  "rPr",
  "sectPr",
  "pPrChange",
]);

const packageWithParagraphProperties = async (propertiesXml: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (!documentXml) {
    panic("The generated package has no main document part.");
  }
  zip.file(
    "word/document.xml",
    documentXml.replace(
      /<w:body>[\s\S]*<\/w:body>/u,
      `<w:body><w:p w14:paraId="12345678"><w:pPr>${propertiesXml}</w:pPr>` +
        "<w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body>",
    ),
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

const NUMBERING_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const NUMBERING_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

/** The generated package carries no numbering part, so the fixture declares one. */
const packageWithLevelProperties = async (propertiesXml: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (!types || !rels) {
    panic("The generated package lost its packaging parts.");
  }
  zip.file(
    "word/numbering.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="900"><w:lvl w:ilvl="0">' +
      '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
      `<w:lvlJc w:val="left"/><w:pPr>${propertiesXml}</w:pPr>` +
      "</w:lvl></w:abstractNum>" +
      '<w:num w:numId="900"><w:abstractNumId w:val="900"/></w:num></w:numbering>',
  );
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING_CONTENT_TYPE}"/></Types>`,
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `<Relationship Id="rIdLevelProperties" Type="${NUMBERING_RELATIONSHIP}" Target="numbering.xml"/></Relationships>`,
    ),
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

/**
 * An edit that makes the whole-element replay stale, so the writer rebuilds.
 *
 * `runInWithNext` is the one modelled field whose markup is not a `w:pPr`
 * child of its own — it is `<w:specVanish/>` inside the paragraph mark — so
 * the forcing edit cannot collide with any sample the table below states.
 */
const FORCING_EDIT_EVIDENCE = "<w:specVanish/>";

const withEditedFormatting = (parsed: Document): Document => {
  const paragraph = parsed.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    panic("The fixture has no paragraph.");
  }
  paragraph.formatting = { ...paragraph.formatting, runInWithNext: true };
  return parsed;
};

const savedDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const { documentXml } = await unzipDocx(buffer);
  if (typeof documentXml !== "string") {
    panic("The package has no main document part.");
  }
  return documentXml;
};

const packageSavedAfterEdit = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const parsed = await parseDocx(buffer, { preloadFonts: false });
  return await repackDocx(withEditedFormatting(parsed), { updateModifiedDate: false });
};

const savedAfterEdit = async (propertiesXml: string): Promise<string> =>
  await savedDocumentXml(
    await packageSavedAfterEdit(await packageWithParagraphProperties(propertiesXml)),
  );

const savedThroughEditor = async (propertiesXml: string): Promise<string> => {
  const parsed = await parseDocx(await packageWithParagraphProperties(propertiesXml), {
    preloadFonts: false,
  });
  const projected = fromProseDoc(toProseDoc(parsed), parsed);
  return await savedDocumentXml(
    await repackDocx(withEditedFormatting(projected), { updateModifiedDate: false }),
  );
};

describe("a paragraph property set keeps every declared child", () => {
  test("the sample table is total over the generated declared-child set", () => {
    expect(Object.keys(DECLARED_CHILD_SAMPLES).toSorted()).toEqual([...CHILD_NAMES].toSorted());
  });

  test(
    "every declared child survives a save that cannot replay the element",
    async () => {
      for (const child of CHILD_NAMES) {
        const saved = await savedAfterEdit(DECLARED_CHILD_SAMPLES[child]);
        expect(saved).toContain(FORCING_EDIT_EVIDENCE);
        expect(saved, `w:${child} was lost by a capture-free save`).toContain(evidenceOf(child));
      }
    },
    ONE_PACKAGE_PER_DECLARED_CHILD_MS,
  );

  test(
    "every declared child survives the save after it",
    async () => {
      for (const child of CHILD_NAMES) {
        const once = await packageSavedAfterEdit(
          await packageWithParagraphProperties(DECLARED_CHILD_SAMPLES[child]),
        );
        const twice = await savedDocumentXml(await packageSavedAfterEdit(once));
        expect(twice, `w:${child} is not a fixed point`).toContain(evidenceOf(child));
      }
    },
    ONE_PACKAGE_PER_DECLARED_CHILD_MS,
  );

  test(
    "every declared child survives the editor projection",
    async () => {
      for (const child of CHILD_NAMES) {
        const saved = await savedThroughEditor(DECLARED_CHILD_SAMPLES[child]);
        expect(saved, `w:${child} was lost by the editor projection`).toContain(evidenceOf(child));
      }
    },
    ONE_PACKAGE_PER_DECLARED_CHILD_MS,
  );

  test("a captured child comes back at the ordinal the schema gives its name", async () => {
    // `w:cnfStyle` is declared last of the thirty-three and folio models none
    // of it; `w:jc` is declared second to last and folio models it. A sink
    // that counted modelled siblings rather than schema ordinals would write
    // the capture ahead of the property, which is markup Word refuses.
    const saved = await savedAfterEdit(
      `${DECLARED_CHILD_SAMPLES.cnfStyle}${DECLARED_CHILD_SAMPLES.jc}`,
    );
    expect(saved.indexOf("<w:cnfStyle ")).toBeGreaterThan(saved.indexOf("<w:jc "));
  });

  test("an undeclared child comes back beside the neighbour it was authored after", async () => {
    const vendor = '<x:note xmlns:x="urn:example:vendor" x:kind="aside">kept</x:note>';
    const saved = await savedAfterEdit(`${DECLARED_CHILD_SAMPLES.jc}${vendor}`);
    expect(saved).toContain("x:kind=");
    expect(saved.indexOf("x:kind=")).toBeGreaterThan(saved.indexOf("<w:jc "));
  });

  test("a property stated twice keeps the value it stated first and the bytes of the repeat", async () => {
    const saved = await savedAfterEdit('<w:jc w:val="center"/><w:jc w:val="both"/>');
    expect(saved).toContain('<w:jc w:val="center"/>');
    expect(saved).toContain('<w:jc w:val="both"/>');
  });

  test("a property whose value the reader refuses keeps its bytes", () => {
    // Nothing in the model can hold `w:val="madeUp"`, and the reader's
    // enumeration is not a set a handler map keyed by name can state. The
    // handler answers that it took nothing and the element is kept.
    fc.assert(
      fc.asyncProperty(fc.constantFrom("madeUp", "", "LEFT"), async (value) => {
        const saved = await savedAfterEdit(`<w:jc w:val="${value}"/>`);
        expect(saved).toContain(`<w:jc w:val="${value}"/>`);
      }),
      propertyConfig({ numRuns: 3 }),
    );
  });

  /**
   * A repack copies `word/styles.xml` through byte for byte, so a save-leg
   * assertion there would pass on the strength of a file copy. The law for the
   * style tier is the writer: what the one `<w:pPr>` writer produces from the
   * style's parsed property set is what a rebuilt styles part would carry.
   */
  test(
    "a style's property set keeps the same children through the one writer",
    async () => {
      for (const child of CHILD_NAMES) {
        if (CHILDREN_OUTSIDE_THE_GENERAL_SET.has(child)) {
          continue;
        }
        const zip = await JSZip.loadAsync(await createEmptyDocx());
        const stylesXml = await zip.file("word/styles.xml")?.async("text");
        if (!stylesXml) {
          panic("The generated package has no styles part.");
        }
        zip.file(
          "word/styles.xml",
          stylesXml.replace(
            "</w:styles>",
            '<w:style w:type="paragraph" w:styleId="Sample"><w:name w:val="Sample"/>' +
              `<w:pPr>${DECLARED_CHILD_SAMPLES[child]}</w:pPr></w:style></w:styles>`,
          ),
        );
        const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
          preloadFonts: false,
        });
        const style = parsed.package.styles?.styles.find(({ styleId }) => styleId === "Sample");
        expect(style, `the Sample style carrying w:${child} did not parse`).toBeDefined();
        expect(
          serializeParagraphFormatting(style?.pPr),
          `w:${child} was lost by the style tier`,
        ).toContain(evidenceOf(child));
      }
    },
    ONE_PACKAGE_PER_DECLARED_CHILD_MS,
  );

  /**
   * The fourth owner. A numbering level read its own `w:pPr` with a private
   * copy of the reader that took an indent and a tab list and let the other
   * thirty-one declared children fall off the end of the walk, while the
   * level's writer was already the shared one — so every child it did not
   * read was written back as nothing. `word/numbering.xml` is another part a
   * repack copies through, so the law is the writer here for the reason the
   * style tier's is.
   */
  test(
    "a numbering level's property set keeps the same children through the one writer",
    async () => {
      for (const child of CHILD_NAMES) {
        if (CHILDREN_OUTSIDE_THE_GENERAL_SET.has(child)) {
          continue;
        }
        const parsed = await parseDocx(
          await packageWithLevelProperties(DECLARED_CHILD_SAMPLES[child]),
          {
            preloadFonts: false,
          },
        );
        const level = parsed.package.numbering?.abstractNums.at(0)?.levels.at(0);
        expect(level, `the level carrying w:${child} did not parse`).toBeDefined();
        expect(
          serializeParagraphFormatting(level?.pPr),
          `w:${child} was lost by the numbering tier`,
        ).toContain(evidenceOf(child));
      }
    },
    ONE_PACKAGE_PER_DECLARED_CHILD_MS,
  );
});
