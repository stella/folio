/**
 * A run's single and double strike survive an edit and save whatever its
 * paragraph style states for either. A double strike paints over a single
 * one, so a run that states `w:strike` under a style's `w:dstrike` shows no
 * single strike in the editor; the save still writes what the run stated, and
 * the effective strike and double strike of every saved run are the source's.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { RELATIONSHIP_TYPES } from "./relsParser";

setDefaultTimeout(propertyTestTimeout(120_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const TOGGLES = ["strike", "dstrike"] as const;
type Toggle = (typeof TOGGLES)[number];
type Stated = Partial<Record<Toggle, boolean>>;

const statedArbitrary: fc.Arbitrary<Stated> = fc.record(
  { strike: fc.boolean(), dstrike: fc.boolean() },
  { requiredKeys: [] },
);

const rPr = (stated: Stated) =>
  TOGGLES.map((toggle) => {
    const value = stated[toggle];
    if (value === undefined) return "";
    return value ? `<w:${toggle}/>` : `<w:${toggle} w:val="0"/>`;
  }).join("");

const packageOf = async (style: Stated, run: Stated) => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/></Relationships>`,
  );
  zip.set(
    "word/styles.xml",
    `${HEAD}<w:styles xmlns:w="${W}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/><w:rPr>${rPr(style)}</w:rPr></w:style>` +
      `</w:styles>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="11111111"><w:pPr><w:pStyle w:val="Body"/></w:pPr>` +
      `<w:r><w:rPr>${rPr(run)}</w:rPr><w:t xml:space="preserve">Kept words edited</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

/** What each saved run shows, its own statements over the style's. */
const effective = (xml: string, style: Stated) =>
  [...xml.matchAll(/<w:r>(.*?)<\/w:r>/gu)].map(([, run = ""]) =>
    Object.fromEntries(
      TOGGLES.map((toggle) => {
        const stated = new RegExp(
          `<w:${toggle}( w:val="([^"]*)")?/>`,
          "u",
        ).exec(run);
        const value =
          stated === null
            ? (style[toggle] ?? false)
            : !["0", "false", "off"].includes(stated[2] ?? "true");
        return [toggle, value];
      }),
    ),
  );

describe("strike under a paragraph style", () => {
  test("an edited run keeps the strike and double strike it showed", async () => {
    await assertProperty(
      fc.asyncProperty(statedArbitrary, statedArbitrary, async (style, run) => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await packageOf(style, run),
        );
        const applied = reviewer.applyOperations(
          [
            {
              id: "edit",
              type: "replaceInBlock",
              blockId: "11111111",
              find: "edited",
              replace: "changed",
            },
          ],
          { mode: "direct", undefinedReferences: "refuse" },
        );
        expect(applied.applied).toHaveLength(1);
        const saved = await ZipArchive.read(await reviewer.toBuffer());
        const xml = (await saved.entry("word/document.xml")?.text()) ?? "";
        const source = effective(`<w:r>${rPr(run)}</w:r>`, style)[0];
        const runs = effective(xml, style);
        expect(runs.length).toBeGreaterThan(0);
        for (const shown of runs) expect(shown).toEqual(source);
      }),
      { id: "strike under a paragraph style survives an edit", numRuns: 60 },
    );
  });
});
