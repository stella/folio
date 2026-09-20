/**
 * An element keeps the attributes folio has no field for.
 *
 * Word writes a revision-session id on nearly every paragraph, run, row and
 * section — `w:rsidR` and its family — and folio rebuilt every one of those
 * elements without it, so opening a document and saving it rewrote the whole
 * revision history. The remainder puts them back on the record they were
 * authored on.
 *
 * The property runs over arbitrary subsets rather than one fixture because the
 * defect is per attribute: a writer that emits the family it knows about keeps
 * every example anybody thought to write down and still loses the one it did
 * not. The editor assertions are the boundary the design turns on — a record
 * the editor creates must not inherit a revision id from anywhere.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { fromProseDoc, proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** The `w:rsid*` attributes each owner's complex type declares. */
const RSID_ATTRIBUTES = {
  p: ["rsidR", "rsidRPr", "rsidDel", "rsidP", "rsidRDefault"],
  r: ["rsidR", "rsidRPr", "rsidDel"],
  tr: ["rsidR", "rsidRPr", "rsidDel", "rsidTr"],
  sectPr: ["rsidR", "rsidRPr", "rsidDel", "rsidSect"],
} as const satisfies Record<string, readonly string[]>;

type Owner = keyof typeof RSID_ATTRIBUTES;

/** `ST_LongHexNumber`: exactly eight hex digits, one distinct value per slot. */
const valueFor = (owner: Owner, attribute: string): string =>
  `00${(((owner.length * 31 + attribute.length) * 2654435761) % 0xff_ff_ff).toString(16).padStart(6, "0").toUpperCase()}`;

const attributesFor = (owner: Owner, chosen: readonly string[]): string =>
  chosen.map((name) => ` w:${name}="${valueFor(owner, name)}"`).join("");

type Chosen = Readonly<Record<Owner, readonly string[]>>;

const documentXml = (chosen: Chosen): string =>
  `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body>` +
  `<w:p${attributesFor("p", chosen.p)}><w:r${attributesFor("r", chosen.r)}><w:t>folio</w:t></w:r></w:p>` +
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>` +
  `<w:tr${attributesFor("tr", chosen.tr)}><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>` +
  `<w:sectPr${attributesFor("sectPr", chosen.sectPr)}><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>` +
  "</w:body></w:document>";

const packageFor = async (xml: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", xml);
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("text")) ?? "";

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

const open = (xml: string): Promise<Document> =>
  packageFor(xml).then((buffer) => parseDocx(buffer, { preloadFonts: false }));

/** The start tag of the first `<w:name …>` in the part, attributes included. */
const startTag = (xml: string, name: string): string =>
  new RegExp(`<w:${name}(\\s[^>]*)?/?>`, "u").exec(xml)?.[0] ?? "";

const expectCarried = (xml: string, chosen: Chosen, owners: readonly Owner[]): void => {
  for (const owner of owners) {
    const tag = startTag(xml, owner);
    for (const attribute of chosen[owner]) {
      expect(tag).toContain(`w:${attribute}="${valueFor(owner, attribute)}"`);
    }
  }
};

const subsetsOf = (owner: Owner) =>
  fc.subarray([...RSID_ATTRIBUTES[owner]] as string[], { minLength: 0 });

const chosenArbitrary = fc.record({
  p: subsetsOf("p"),
  r: subsetsOf("r"),
  tr: subsetsOf("tr"),
  sectPr: subsetsOf("sectPr"),
});

const ALL_OWNERS = ["p", "r", "tr", "sectPr"] as const;
/** The three owners whose record survives the editor; a run has no record there. */
const EDITOR_OWNERS = ["p", "tr", "sectPr"] as const;

describe("the attribute remainder survives a save", () => {
  test("every subset of the rsid family comes back on the element it was written on", async () => {
    await fc.assert(
      fc.asyncProperty(chosenArbitrary, async (chosen) => {
        const saved = await documentPartOf(await save(await open(documentXml(chosen))));
        expectCarried(saved, chosen, ALL_OWNERS);

        // Save, reopen, save: the second save is where a remainder that only
        // replays and does not re-parse stops being a fixed point.
        expect(await documentPartOf(await save(await open(saved)))).toBe(saved);
      }),
      propertyConfig({ numRuns: 25 }),
    );
  }, 120_000);

  test("a second prefix bound to the same namespace is read once, not kept twice", async () => {
    const xml =
      `${XML_DECLARATION}<w:document xmlns:w="${W}" xmlns:altw="${W}" ` +
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>' +
      '<w:p altw:rsidR="00ABCDEF" altw:paraId="1F2E3D4C" w14:textId="0A0B0C0D">' +
      "<w:r><w:t>folio</w:t></w:r></w:p></w:body></w:document>";
    const tag = startTag(await documentPartOf(await save(await open(xml))), "p");

    // The remainder is resolved, so the alternative binding comes back under
    // the prefix the rebuilt part declares for that namespace.
    expect(tag).toContain('w:rsidR="00ABCDEF"');
    // `paraId` is modelled whatever prefix spelled it, so the remainder must
    // not write a second copy of the id the paragraph record already holds.
    expect([...tag.matchAll(/paraId=/gu)]).toHaveLength(1);
    expect([...tag.matchAll(/textId=/gu)]).toHaveLength(1);
  });
});

describe("the attribute remainder follows the record through the editor", () => {
  test("an authored element's remainder survives the projection unchanged", async () => {
    await fc.assert(
      fc.asyncProperty(chosenArbitrary, async (chosen) => {
        const parsed = await open(documentXml(chosen));
        const projected = fromProseDoc(toProseDoc(parsed), parsed);
        expectCarried(await documentPartOf(await save(projected)), chosen, EDITOR_OWNERS);
      }),
      propertyConfig({ numRuns: 25 }),
    );
  }, 120_000);

  test("a paragraph the editor creates from scratch has no remainder", async () => {
    const parsed = await open(documentXml({ p: ["rsidR"], r: [], tr: [], sectPr: [] }));
    const projected = toProseDoc(parsed);
    const created = schema.nodes["paragraph"]?.createAndFill();
    if (!created) {
      throw new Error("the schema refused an empty paragraph");
    }

    const rebuilt = fromProseDoc(
      schema.node("doc", projected.attrs, [created, ...projected.content.content]),
      parsed,
    );
    const first = rebuilt.package.document.content.at(0) as Paragraph;
    expect(first.type).toBe("paragraph");
    expect(first.preservedAttributes).toBeUndefined();
  });

  test("a paragraph that only hosts an anchored drawing keeps its remainder", async () => {
    // Word writes a floating shape into a paragraph of its own. folio lifts
    // the shape out as a block node and drops that paragraph from the
    // projection, so the node it leaves behind is the only carrier the host's
    // attributes have.
    const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    const xml =
      `${XML_DECLARATION}<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" ` +
      `xmlns:wps="${WPS}"><w:body>` +
      '<w:p w:rsidR="00ABCDEF"><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" ' +
      'distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" ' +
      'layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="1000000" cy="500000"/><wp:docPr id="1" name="Text Box 1"/>' +
      `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:spPr/><wps:txbx>` +
      "<w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent>" +
      "</wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>" +
      "</wp:anchor></w:drawing></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
      "</w:body></w:document>";

    const parsed = await open(xml);
    const host = parsed.package.document.content.at(0) as Paragraph;
    expect(host.preservedAttributes).toBeDefined();

    const projected = fromProseDoc(toProseDoc(parsed), parsed);
    const rebuilt = projected.package.document.content.at(0) as Paragraph;
    expect(rebuilt.preservedAttributes).toEqual(host.preservedAttributes);
    expect(startTag(await documentPartOf(await save(projected)), "p")).toContain(
      'w:rsidR="00ABCDEF"',
    );
  });

  test("a split paragraph keeps the remainder on the half that was authored", async () => {
    const parsed = await open(documentXml({ p: ["rsidR"], r: [], tr: [], sectPr: [] }));
    const projection = toProseDoc(parsed);
    const authored = projection.content.content.at(0);
    if (!authored) {
      throw new Error("the projection produced no paragraph");
    }

    // What `splitBlock` produces: two nodes over one attrs object, so both
    // halves hold the very array the authored paragraph carried.
    const halves = [
      authored.type.create(authored.attrs, authored.content),
      authored.type.create(authored.attrs, authored.content),
    ];
    // `proseDocToBlocks` rather than `fromProseDoc`: the paragraph-property
    // source contract refuses two paragraphs carrying one source token, which
    // is a different rule about the same copy and would mask this one.
    const blocks = proseDocToBlocks(
      schema.node("doc", projection.attrs, halves),
      parsed.package.document.content,
      parsed.package.styles,
    );
    const [head, tail] = blocks as Paragraph[];

    expect(head?.preservedAttributes).toBeDefined();
    expect(tail?.preservedAttributes).toBeUndefined();
  });
});
