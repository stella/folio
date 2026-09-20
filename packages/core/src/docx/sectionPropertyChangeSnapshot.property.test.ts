/**
 * A `w:sectPrChange` holds a whole section, and a rebuild has to write it back.
 *
 * The snapshot's type is `CT_SectPrBase`: `CT_SectPr` without the header and
 * footer references and without a `w:sectPrChange` of its own. Everything else
 * it declares is declared on the live section too, so the snapshot is read by
 * the same `parseSectionProperties` and written by the same
 * `serializeSectionProperties` — the property below is what says so, because a
 * reader narrowed to a private subset would pass every test that only ever
 * looks at the live section.
 *
 * The declared list is read from the committed schema graph rather than
 * restated here, so a schema refresh that adds a child to `CT_SectPrBase` fails
 * this file until somebody says what the snapshot does with it.
 */

import { readFileSync } from "node:fs";

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type SchemaChild = {
  kind?: string;
  name?: string;
  owner?: string;
  ref?: string;
};

/** Every element `CT_SectPrBase` declares, groups expanded, in schema order. */
const DECLARED_SNAPSHOT_CHILDREN: readonly string[] = (() => {
  const graph = JSON.parse(
    readFileSync(
      new URL(
        "../../../../specifications/generated/docx-transitional-schema.gen.json",
        import.meta.url,
      ),
      "utf8",
    ),
  ) as { children: SchemaChild[] };
  const byOwner = new Map<string, SchemaChild[]>();
  for (const child of graph.children) {
    const owner = child.owner;
    if (owner === undefined) {
      continue;
    }
    const declared = byOwner.get(owner);
    if (declared === undefined) {
      byOwner.set(owner, [child]);
      continue;
    }
    declared.push(child);
  }
  const walk = (owner: string, seen: ReadonlySet<string>): string[] =>
    (byOwner.get(owner) ?? []).flatMap((child) => {
      if (child.kind === "element") {
        return child.name === undefined ? [] : [child.name];
      }
      const group = child.ref === undefined ? undefined : `group:${child.ref}`;
      return group === undefined || seen.has(group) ? [] : walk(group, new Set([...seen, group]));
    });
  const declared = walk(`complexType:{${WORD_NAMESPACE}}CT_SectPrBase`, new Set());
  if (declared.length === 0) {
    throw new Error("CT_SectPrBase is missing from the schema graph");
  }
  return declared;
})();

/**
 * One authored spelling per declared child.
 *
 * Each states something of its own, which is the form a document carries: an
 * element that states nothing is a separate question, and the two the contract
 * records as lost bare are the subject of the test below this one.
 */
const AUTHORED_CHILD = {
  footnotePr: '<w:footnotePr><w:numFmt w:val="lowerRoman"/></w:footnotePr>',
  cols: '<w:cols w:num="2" w:space="425"/>',
  formProt: '<w:formProt w:val="1"/>',
  vAlign: '<w:vAlign w:val="center"/>',
  noEndnote: '<w:noEndnote w:val="1"/>',
  titlePg: '<w:titlePg w:val="1"/>',
  textDirection: '<w:textDirection w:val="lrTb"/>',
  bidi: '<w:bidi w:val="1"/>',
  rtlGutter: '<w:rtlGutter w:val="1"/>',
  docGrid: '<w:docGrid w:type="lines" w:linePitch="360"/>',
  printerSettings: '<w:printerSettings r:id="rIdPrinter"/>',
  endnotePr: '<w:endnotePr><w:numFmt w:val="lowerLetter"/></w:endnotePr>',
  type: '<w:type w:val="nextPage"/>',
  pgSz: '<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>',
  pgMar: '<w:pgMar w:top="1417" w:right="1134" w:bottom="1417" w:left="1134"/>',
  paperSrc: '<w:paperSrc w:first="7" w:other="8"/>',
  pgBorders:
    '<w:pgBorders w:display="allPages"><w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/></w:pgBorders>',
  lnNumType: '<w:lnNumType w:countBy="5" w:start="1" w:restart="newPage"/>',
  pgNumType: '<w:pgNumType w:start="3" w:fmt="decimal"/>',
} as const;

type DeclaredChild = keyof typeof AUTHORED_CHILD;

const DECLARED_CHILDREN = Object.keys(AUTHORED_CHILD) as DeclaredChild[];

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const documentWith = (snapshotChildren: string): string =>
  `${XML_DECLARATION}<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${R_NAMESPACE}"><w:body>` +
  "<w:p><w:pPr><w:sectPr>" +
  '<w:sectPrChange w:id="1" w:author="Reviewer" w:date="2024-01-02T03:04:00Z">' +
  `<w:sectPr>${snapshotChildren}</w:sectPr>` +
  "</w:sectPrChange>" +
  "</w:sectPr></w:pPr><w:r><w:t>folio</w:t></w:r></w:p>" +
  "<w:sectPr/></w:body></w:document>";

const savedSnapshot = async (snapshotChildren: string): Promise<string> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", documentWith(snapshotChildren));
  const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
    preloadFonts: false,
  });
  const saved = await repackDocx(parsed, { updateModifiedDate: false });
  const part =
    (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
  // The snapshot only, because a live section carrying the same child would
  // answer the question the fixture did not ask.
  return /<w:sectPrChange\b.*?<\/w:sectPrChange>/su.exec(part)?.[0] ?? "";
};

describe("a w:sectPrChange keeps the section it holds", () => {
  test("the authored spellings cover every child CT_SectPrBase declares", () => {
    expect([...DECLARED_CHILDREN].sort()).toEqual([...DECLARED_SNAPSHOT_CHILDREN].sort());
  });

  test("every declared child comes back inside the snapshot", async () => {
    await fc.assert(
      fc.asyncProperty(fc.constantFrom(...DECLARED_CHILDREN), async (name) => {
        const snapshot = await savedSnapshot(AUTHORED_CHILD[name]);
        expect({ name, kept: snapshot.includes(`<w:${name}`) }).toEqual({ name, kept: true });
      }),
      propertyConfig({ numRuns: DECLARED_CHILDREN.length }),
    );
  });

  /**
   * `w:lnNumType` and `w:pgBorders` declare every attribute and child optional,
   * so the element alone is a legal document that states a section is line
   * numbered, or bordered, with the defaults. Both serializers dropped it, and
   * inside a snapshot that loss is silent: the change record is written with an
   * empty `<w:sectPr/>` rather than refused.
   */
  test("an element that states nothing is still an element the source had", async () => {
    await fc.assert(
      fc.asyncProperty(fc.constantFrom("lnNumType", "pgBorders"), async (name) => {
        const snapshot = await savedSnapshot(`<w:${name}/>`);
        expect({ name, kept: snapshot.includes(`<w:${name}`) }).toEqual({ name, kept: true });
      }),
      propertyConfig({ numRuns: 2 }),
    );
  });
});
