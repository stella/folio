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
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { EditorState, NodeSelection, TextSelection } from "prosemirror-state";

import { propertyConfig } from "../../../../test/property-testing";

import { fromProseDoc, proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { BaseKeymapExtension } from "../prosemirror/extensions/features/BaseKeymapExtension";
import { schema } from "../prosemirror/schema";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

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

// ============================================================================
// EVERY BLOCK CONTAINER
// ============================================================================

/**
 * The rule holds per block container or it does not hold at all.
 *
 * A record is in a text box, a cell, an `w:sdt` or the body, and the editor
 * reaches all four the same way, so the census below has to come out the same
 * wherever the fixture puts the record. The containers are read off the editor
 * schema rather than listed here: a node that gains paragraph content has to
 * be given a fixture before this file passes again, which is what keeps the
 * list from drifting behind the schema.
 */
const paragraphContainersInSchema = (): string[] => {
  const paragraph = schema.nodes["paragraph"];
  if (!paragraph) {
    throw new Error("the schema declares no paragraph");
  }
  return Object.values(schema.nodes)
    .filter((type) => type.contentMatch.matchType(paragraph) !== null)
    .map(({ name }) => name);
};

const TABLE_GRID = '<w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid>';

const tableAround = (inner: string, look: string): string =>
  `<w:tbl><w:tblPr>${look}</w:tblPr>${TABLE_GRID}` +
  `<w:tr><w:tc><w:tcPr/>${inner}</w:tc></w:tr></w:tbl>`;

/**
 * A floating shape, which is how `w:txbxContent` is reached: through a run
 * rather than through a block child.
 */
const anchoredTextBoxAround = (inner: string): string =>
  '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
  'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1000000" cy="500000"/><wp:docPr id="1" name="Text Box 1"/>' +
  `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:spPr/><wps:txbx>` +
  `<w:txbxContent>${inner}</w:txbxContent>` +
  "</wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>" +
  "</wp:anchor></w:drawing></w:r></w:p>";

const CONTAINER_FIXTURES: Record<string, (inner: string) => string> = {
  doc: (inner) => inner,
  blockSdt: (inner) =>
    `<w:sdt><w:sdtPr><w:tag w:val="container"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`,
  tableCell: (inner) => tableAround(inner, ""),
  // `@w:firstRow` is what makes the projection emit `tableHeader` rather than
  // a second `tableCell`.
  tableHeader: (inner) => tableAround(inner, '<w:tblLook w:firstRow="1"/>'),
  textBox: anchoredTextBoxAround,
};

/** One paragraph record and one row record, each with its own rsid subset. */
const recordsXml = (chosen: Pick<Chosen, "p" | "tr">): string =>
  `<w:p${attributesFor("p", chosen.p)}><w:r><w:t>folio</w:t></w:r></w:p>` +
  `<w:tbl><w:tblPr/>${TABLE_GRID}` +
  `<w:tr${attributesFor("tr", chosen.tr)}><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>`;

const containerDocumentXml = (body: string): string =>
  `${XML_DECLARATION}<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" ` +
  `xmlns:wps="${WPS}"><w:body>${body}` +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';

/**
 * Every remainder in a block tree, found by walking the data rather than by
 * asking the traversal under test where the records are.
 */
const remainderCensus = (value: unknown, path = ""): string[] => {
  if (Array.isArray(value)) {
    return value.flatMap((item) => remainderCensus(item, path));
  }
  if (value === null || typeof value !== "object") {
    return [];
  }
  const found: string[] = [];
  for (const [key, entry] of Object.entries(value)) {
    if (key === "preservedAttributes" && Array.isArray(entry)) {
      for (const attribute of entry as { name: string; value: string }[]) {
        found.push(`${path}${attribute.name}=${attribute.value}`);
      }
      continue;
    }
    found.push(...remainderCensus(entry, key === "textBody" ? `${path}textBox/` : path));
  }
  return found.sort();
};

const nodeTypeNames = (doc: PMNode): Set<string> => {
  const names = new Set<string>([doc.type.name]);
  doc.descendants((node) => {
    names.add(node.type.name);
    return true;
  });
  return names;
};

/**
 * The attrs a copy holds by reference rather than by value, which is what
 * tells a copy apart from a record that parsed its own.
 */
const SHARED_CARRIER_ATTRS = ["_preservedAttributes", "_sectionProperties"] as const;

/** What `splitBlock` leaves behind: `copies` extra nodes over one attrs object. */
const duplicateRecordNodes = (node: PMNode, copies = 1): PMNode => {
  if (node.childCount === 0) {
    return node;
  }
  const children: PMNode[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    const mapped = duplicateRecordNodes(child, copies);
    children.push(mapped);
    if (
      (child.type.name === "paragraph" || child.type.name === "tableRow") &&
      SHARED_CARRIER_ATTRS.some((attr) => child.attrs[attr])
    ) {
      for (let copy = 0; copy < copies; copy++) {
        children.push(child.type.create(child.attrs, child.content, child.marks));
      }
    }
  });
  return node.type.create(node.attrs, Fragment.fromArray(children), node.marks);
};

const containerArbitrary = fc.record({ p: subsetsOf("p"), tr: subsetsOf("tr") });

describe("the attribute remainder holds in every block container", () => {
  test("every container the schema declares has a fixture", () => {
    expect([...paragraphContainersInSchema()].sort()).toEqual(
      Object.keys(CONTAINER_FIXTURES).sort(),
    );
  });

  for (const [container, wrap] of Object.entries(CONTAINER_FIXTURES)) {
    test(`${container}: a record inside it keeps the remainder it was authored with`, async () => {
      await fc.assert(
        fc.asyncProperty(containerArbitrary, async (chosen) => {
          const parsed = await open(containerDocumentXml(wrap(recordsXml(chosen))));
          const authored = remainderCensus(parsed.package.document.content);
          const projection = toProseDoc(parsed);

          // A fixture that stopped producing the container would otherwise
          // assert about the body and pass.
          expect(nodeTypeNames(projection)).toContain(container);

          const rebuilt = fromProseDoc(projection, parsed);
          expect(remainderCensus(rebuilt.package.document.content)).toEqual(authored);
        }),
        propertyConfig({ numRuns: 10 }),
      );
    }, 120_000);

    test(`${container}: a copy the editor made inside it claims no revision session`, async () => {
      await fc.assert(
        fc.asyncProperty(containerArbitrary, async (chosen) => {
          const parsed = await open(containerDocumentXml(wrap(recordsXml(chosen))));
          const authored = remainderCensus(parsed.package.document.content);

          // `proseDocToBlocks` rather than `fromProseDoc`: the
          // paragraph-property source contract refuses two paragraphs carrying
          // one source token, which is a different rule about the same copy.
          const blocks: BlockContent[] = proseDocToBlocks(
            duplicateRecordNodes(toProseDoc(parsed)),
            parsed.package.document.content,
            parsed.package.styles,
          );
          expect(remainderCensus(blocks)).toEqual(authored);
        }),
        propertyConfig({ numRuns: 10 }),
      );
    }, 120_000);
  }
});

// ============================================================================
// THE SECTION CARRIER
// ============================================================================

/**
 * The section break is the copy rule read the other way round.
 *
 * `w:sectPr` inside a `w:pPr` says the section ends at *this* paragraph's
 * mark, so when a split hands the same properties to two halves the trailing
 * one is the half that still ends the section — the opposite of the remainder,
 * which stays with the half that was authored. Writing the break on both ends
 * the section twice: a section nobody added, whose `w:sectPr` repeats the
 * `w:rsidSect` of the real one and so claims its revision history as well.
 *
 * The count is asserted against the same fixture's own save rather than a
 * literal, so a fixture that gains a section does not need the number edited
 * in two places.
 */

const SECTION_FIXTURE_TEXT = "folio";

const sectionDocumentXml = ({ p, sectPr }: Pick<Chosen, "p" | "sectPr">): string =>
  `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body>` +
  "<w:p><w:r><w:t>lead</w:t></w:r></w:p>" +
  `<w:p${attributesFor("p", p)}><w:pPr><w:sectPr${attributesFor("sectPr", sectPr)}>` +
  '<w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>' +
  `<w:r><w:t>${SECTION_FIXTURE_TEXT}</w:t></w:r></w:p>` +
  "<w:p><w:r><w:t>tail</w:t></w:r></w:p>" +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
  "</w:body></w:document>";

/** The paragraph the fixture's section ends at, before any editing. */
const SECTION_CARRIER_INDEX = 1;

const SECT_PR = /<w:sectPr[\s/>]/gu;

const occurrences = (xml: string, pattern: RegExp): number => [...xml.matchAll(pattern)].length;

const valuePattern = (owner: Owner, attribute: string): RegExp =>
  new RegExp(`w:${attribute}="${valueFor(owner, attribute)}"`, "gu");

const documentWithContent = (parsed: Document, content: BlockContent[]): Document => ({
  ...parsed,
  package: {
    ...parsed.package,
    document: { ...parsed.package.document, content },
  },
});

const sectionCarrierIndexes = (blocks: readonly BlockContent[]): number[] =>
  blocks.flatMap((block, index) =>
    block.type === "paragraph" && block.sectionProperties ? [index] : [],
  );

const sectionArbitrary = fc.record({
  p: subsetsOf("p"),
  sectPr: subsetsOf("sectPr"),
  copies: fc.integer({ min: 1, max: 4 }),
});

describe("the section break belongs to the last paragraph of its section", () => {
  test("splitting the section-ending paragraph writes one w:sectPr, on the last half", async () => {
    await fc.assert(
      fc.asyncProperty(sectionArbitrary, async ({ copies, ...chosen }) => {
        const parsed = await open(sectionDocumentXml(chosen));
        const authored = await documentPartOf(await save(parsed));

        const blocks = proseDocToBlocks(
          duplicateRecordNodes(toProseDoc(parsed), copies),
          parsed.package.document.content,
          parsed.package.styles,
        );
        expect(sectionCarrierIndexes(blocks)).toEqual([SECTION_CARRIER_INDEX + copies]);

        const saved = await documentPartOf(await save(documentWithContent(parsed, blocks)));
        expect(occurrences(saved, SECT_PR)).toBe(occurrences(authored, SECT_PR));
        for (const attribute of chosen.sectPr) {
          expect(occurrences(saved, valuePattern("sectPr", attribute))).toBe(1);
        }
        for (const attribute of chosen.p) {
          expect(occurrences(saved, valuePattern("p", attribute))).toBe(1);
        }
      }),
      propertyConfig({ numRuns: 15 }),
    );
  }, 120_000);

  test("joining it with its predecessor keeps the break on the merged paragraph", async () => {
    const chosen = { p: ["rsidR"], sectPr: ["rsidSect"] } as const;
    const parsed = await open(sectionDocumentXml(chosen));
    const authored = await documentPartOf(await save(parsed));
    const projection = toProseDoc(parsed);

    const lead = projection.child(0);
    const carrier = projection.child(SECTION_CARRIER_INDEX);
    expect(carrier.textContent).toBe(SECTION_FIXTURE_TEXT);

    // Backspace with the caret at the start of the section-ending paragraph:
    // its predecessor's mark goes, its own survives, and the two become one
    // paragraph that still ends the section.
    const backspace = BaseKeymapExtension().onSchemaReady({ schema }).keyboardShortcuts?.[
      "Backspace"
    ];
    if (!backspace) {
      throw new Error("the base keymap binds no Backspace");
    }
    const state = EditorState.create({ doc: projection });
    const caretAtCarrier = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, lead.nodeSize + 1)),
    );
    let joined = caretAtCarrier;
    expect(
      backspace(caretAtCarrier, (tr) => {
        joined = caretAtCarrier.apply(tr);
      }),
    ).toBe(true);

    const merged = joined.doc.child(0);
    expect(merged.textContent).toBe(`lead${SECTION_FIXTURE_TEXT}`);
    // The very object, not an equal one: the properties were carried across
    // the join rather than rebuilt from the break type.
    expect(merged.attrs["_sectionProperties"]).toBe(carrier.attrs["_sectionProperties"]);

    const blocks = proseDocToBlocks(
      joined.doc,
      parsed.package.document.content,
      parsed.package.styles,
    );
    expect(sectionCarrierIndexes(blocks)).toEqual([0]);

    const saved = await documentPartOf(await save(documentWithContent(parsed, blocks)));
    expect(occurrences(saved, SECT_PR)).toBe(occurrences(authored, SECT_PR));
    expect(occurrences(saved, valuePattern("sectPr", "rsidSect"))).toBe(1);
  });

  test("deleting the section-ending paragraph outright takes its section with it", async () => {
    // The transfer follows a paragraph mark that survived an edit. A mark that
    // was deleted rather than joined away leaves nothing to inherit it, or a
    // section could never be removed at all. The save leg refuses a repack
    // that drops a section, so the assertion stops at the model.
    const parsed = await open(sectionDocumentXml({ p: [], sectPr: ["rsidSect"] }));
    const projection = toProseDoc(parsed);
    const backspace = BaseKeymapExtension().onSchemaReady({ schema }).keyboardShortcuts?.[
      "Backspace"
    ];
    if (!backspace) {
      throw new Error("the base keymap binds no Backspace");
    }

    const state = EditorState.create({ doc: projection });
    const selected = state.apply(
      state.tr.setSelection(NodeSelection.create(state.doc, projection.child(0).nodeSize)),
    );
    let deleted = selected;
    expect(
      backspace(selected, (tr) => {
        deleted = selected.apply(tr);
      }),
    ).toBe(true);

    const blocks = proseDocToBlocks(
      deleted.doc,
      parsed.package.document.content,
      parsed.package.styles,
    );
    expect(sectionCarrierIndexes(blocks)).toEqual([]);
  });
});
