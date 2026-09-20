/**
 * A property element keeps the attributes the model has no field for.
 *
 * The child dispatcher decides a property element whole: a handler either
 * reads it or hands back its bytes. `<w:ind w:leftChars="100"/>` survives
 * because the reader took nothing from it; `<w:ind w:left="720"
 * w:leftChars="100"/>` — which is what a document actually carries — was
 * modelled, and the character unit went. The container-survival census states
 * one attribute at a time, so it reported the pair as surviving.
 *
 * The property below is the closure of that: an arbitrary subset of the
 * element's *schema* attributes, so every combination of stated-and-modelled
 * and stated-and-not is exercised rather than the one pair a hand-written
 * example would pin. The attribute universe comes from the committed schema
 * graph and the modelled half from `propertyElementAttributes.ts`, so neither
 * is a list this file keeps in step by hand.
 *
 * The assertions are about a capture-free save: the paragraph's formatting is
 * edited first, which is what invalidates the whole-element replay and makes
 * the writers rebuild.
 */

import { readFileSync } from "node:fs";

import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";
import { parseDocx } from "./parser";
import { modelledAttributeNames, PROPERTY_ELEMENT_ATTRIBUTES } from "./propertyElementAttributes";
import { createEmptyDocx, repackDocx } from "./rezip";
import { unzipDocx } from "./unzip";

type SchemaGraph = {
  symbols: Array<{ id?: string; kind?: string; enumValues?: string[] }>;
  attributes: Array<{
    owner?: string;
    name?: string;
    namespace?: string;
    ref?: string;
    type?: string;
    use?: string;
  }>;
  inheritance: Array<{ base?: string; derived?: string }>;
};

const GRAPH: SchemaGraph = JSON.parse(
  readFileSync(
    new URL(
      "../../../../specifications/generated/docx-transitional-schema.gen.json",
      import.meta.url,
    ),
    "utf8",
  ),
) as SchemaGraph;

const WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * A value the attribute's simple type accepts, derived rather than listed.
 *
 * The first enumeration member where the type has one; `1` for the `ST_OnOff`
 * family, which is a union and declares none; a hex triple for a colour; a
 * small integer otherwise. What matters for survival is that folio can parse
 * the value at all, not which of them it is.
 */
const valueForType = (type: string | undefined): string => {
  const symbol = GRAPH.symbols.find(
    (candidate) => candidate.kind === "simpleType" && candidate.id === `simpleType:${type}`,
  );
  const enumerated = symbol?.enumValues?.at(0);
  if (enumerated !== undefined) {
    return enumerated;
  }
  const name = type ?? "";
  if (name.includes("OnOff")) {
    return "1";
  }
  if (name.includes("HexColor")) {
    return "AABBCC";
  }
  return name.includes("HexNumber") ? "80" : "120";
};

type DeclaredAttribute = { name: string; value: string };

/** Every attribute the schema declares on a complex type, in declaration order. */
const declaredAttributesOf = (typeName: string): DeclaredAttribute[] =>
  GRAPH.attributes
    .filter(
      (attribute) =>
        attribute.owner === `complexType:{${WML}}${typeName}` && attribute.namespace === WML,
    )
    .flatMap((attribute) =>
      attribute.name === undefined
        ? []
        : [{ name: attribute.name, value: valueForType(attribute.type) }],
    );

type PropertyElementType = keyof typeof PROPERTY_ELEMENT_ATTRIBUTES;

/**
 * Where each property element sits in a `w:pPr`, so one fixture serves all six.
 *
 * `Record`, never `Partial<Record>`: the key set is the model's own table, so
 * a record that gains an attribute remainder cannot join without a place to
 * be measured in.
 */
const HOSTS = {
  CT_Ind: (attributes: string) => `<w:ind ${attributes}/>`,
  CT_Spacing: (attributes: string) => `<w:spacing ${attributes}/>`,
  CT_FramePr: (attributes: string) => `<w:framePr ${attributes}/>`,
  CT_TabStop: (attributes: string) => `<w:tabs><w:tab ${attributes}/></w:tabs>`,
  CT_Border: (attributes: string) => `<w:pBdr><w:top ${attributes}/></w:pBdr>`,
  CT_Shd: (attributes: string) => `<w:shd ${attributes}/>`,
} as const satisfies Record<PropertyElementType, (attributes: string) => string>;

const PROPERTY_ELEMENT_TYPES = Object.keys(HOSTS) as PropertyElementType[];

/**
 * The attributes the type declares `use="required"`: `w:val` and `w:pos` on a
 * tab stop, `w:val` on a shading. Every subset states them, because an element
 * that omits one is invalid markup rather than a case folio has to survive.
 */
const requiredAttributesOf = (typeName: PropertyElementType): readonly string[] =>
  GRAPH.attributes
    .filter(
      (attribute) =>
        attribute.owner === `complexType:{${WML}}${typeName}` &&
        attribute.namespace === WML &&
        attribute.use === "required",
    )
    .flatMap((attribute) => (attribute.name === undefined ? [] : [attribute.name]));

/**
 * Every attribute the type declares, plus the ones its extensions add.
 *
 * One reader serves a type and everything derived from it, so the modelled
 * set is about the family rather than about the one complex type: folio reads
 * `r:topLeft` off every `CT_Border`, and `CT_PageBorder` is where the schema
 * declares it.
 */
const attributeName = (attribute: SchemaGraph["attributes"][number]): string | undefined =>
  attribute.name ?? attribute.ref?.slice(attribute.ref.indexOf("}") + 1);

const declaredByTypeAndItsExtensions = (typeName: string): SchemaGraph["attributes"] => {
  const owners = new Set([`complexType:{${WML}}${typeName}`]);
  for (let added = true; added;) {
    added = false;
    for (const { base, derived } of GRAPH.inheritance) {
      if (
        base !== undefined &&
        derived !== undefined &&
        owners.has(`complexType:${base}`) &&
        !owners.has(derived)
      ) {
        owners.add(derived);
        added = true;
      }
    }
  }
  return GRAPH.attributes.filter(
    (attribute) => attribute.owner !== undefined && owners.has(attribute.owner),
  );
};

const split = (typeName: PropertyElementType) => {
  const modelled = new Set(modelledAttributeNames(PROPERTY_ELEMENT_ATTRIBUTES[typeName]));
  const declared = declaredAttributesOf(typeName);
  return {
    declared,
    modelled: declared.filter(({ name }) => modelled.has(name)),
    remainder: declared.filter(({ name }) => !modelled.has(name)),
  };
};

const spell = (attributes: readonly DeclaredAttribute[]): string =>
  attributes.map(({ name, value }) => `w:${name}="${value}"`).join(" ");

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

/** An edit the whole-element replay cannot serve, so the writers rebuild. */
const withEditedFormatting = (parsed: Document): Document => {
  const paragraph = parsed.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    panic("The fixture has no paragraph.");
  }
  paragraph.formatting = { ...paragraph.formatting, runInWithNext: true };
  return parsed;
};

const documentXmlOf = async (buffer: ArrayBuffer): Promise<string> => {
  const { documentXml } = await unzipDocx(buffer);
  if (typeof documentXml !== "string") {
    panic("The package has no main document part.");
  }
  return documentXml;
};

const savedAfterEdit = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const parsed = await parseDocx(buffer, { preloadFonts: false });
  return await repackDocx(withEditedFormatting(parsed), { updateModifiedDate: false });
};

const savedThroughEditor = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const parsed = await parseDocx(buffer, { preloadFonts: false });
  const projected = fromProseDoc(toProseDoc(parsed), parsed);
  return await repackDocx(withEditedFormatting(projected), { updateModifiedDate: false });
};

describe("a property element keeps what the model has no field for", () => {
  test("the host table names every type the model states a remainder for", () => {
    expect(PROPERTY_ELEMENT_TYPES.toSorted()).toEqual(
      Object.keys(PROPERTY_ELEMENT_ATTRIBUTES).toSorted(),
    );
  });

  test("every name the model calls modelled is an attribute the type declares", () => {
    // The table is the remainder's predicate, so a name misspelled in it
    // silently moves an attribute into the remainder and the save writes the
    // value twice under two spellings. The schema graph is what says whether
    // the name exists. Names are matched in any namespace, because the
    // remainder decides on the resolved local name, and the extension chain
    // is walked because one reader serves a type and its extensions: folio
    // reads a page border's art relationship ids off every `CT_Border`, and
    // only `CT_PageBorder` declares them.
    for (const typeName of PROPERTY_ELEMENT_TYPES) {
      const declared = new Set(
        declaredByTypeAndItsExtensions(typeName).flatMap((attribute) => {
          const name = attributeName(attribute);
          return name === undefined ? [] : [name];
        }),
      );
      expect(declared.size, `${typeName} is not in the schema graph`).toBeGreaterThan(0);
      for (const name of modelledAttributeNames(PROPERTY_ELEMENT_ATTRIBUTES[typeName])) {
        expect(declared, `${typeName} declares no w:${name}`).toContain(name);
      }
    }
  });

  test(
    "an arbitrary subset survives a save, the save after it, and the editor",
    async () => {
      for (const typeName of PROPERTY_ELEMENT_TYPES) {
        const { modelled, remainder } = split(typeName);
        if (remainder.length === 0) {
          // `CT_TabStop` declares three attributes and the model holds all
          // three, so it has no remainder to lose. It stays in the table so
          // that an attribute a later schema adds lands in one.
          continue;
        }
        const required = requiredAttributesOf(typeName);
        await fc.assert(
          fc.asyncProperty(
            fc.subarray(modelled, { minLength: 1 }),
            fc.subarray(remainder, { minLength: 1 }),
            async (statedModelled, statedRemainder) => {
              const stated = [
                ...modelled.filter(({ name }) => required.includes(name)),
                ...statedModelled.filter(({ name }) => !required.includes(name)),
                ...statedRemainder,
              ];
              const fixture = await packageWithParagraphProperties(HOSTS[typeName](spell(stated)));

              const once = await savedAfterEdit(fixture);
              const twice = await savedAfterEdit(once);
              const editor = await savedThroughEditor(fixture);
              const legs = [
                ["a capture-free save", await documentXmlOf(once)],
                ["the save after it", await documentXmlOf(twice)],
                ["the editor projection", await documentXmlOf(editor)],
              ] as const;

              for (const [leg, saved] of legs) {
                for (const { name, value } of statedRemainder) {
                  expect(saved, `w:${name} on ${typeName} was lost by ${leg}`).toContain(
                    `w:${name}="${value}"`,
                  );
                }
              }
            },
          ),
          propertyConfig({ numRuns: 12 }),
        );
      }
    },
    propertyTestTimeout(120_000),
  );

  test("a character unit stated beside the twip it qualifies is not dropped", async () => {
    // The counterexample the census could not see. `<w:ind w:leftChars="100"/>`
    // alone was kept whole; with `w:left` beside it the element is modelled and
    // the character unit had nowhere to go.
    const saved = await documentXmlOf(
      await savedAfterEdit(
        await packageWithParagraphProperties('<w:ind w:left="720" w:leftChars="100"/>'),
      ),
    );
    expect(saved).toContain('w:left="720"');
    expect(saved).toContain('w:leftChars="100"');
  });

  test("a preserved attribute is written once, under the prefix the part binds", async () => {
    // A second binding of the WordprocessingML namespace spells the same
    // attribute differently; the remainder decides on the resolved local name,
    // so the save writes one `w:leftChars` rather than two spellings of it.
    const saved = await documentXmlOf(
      await savedAfterEdit(
        await packageWithParagraphProperties(
          `<w:ind xmlns:altw="${WML}" w:left="720" altw:leftChars="100"/>`,
        ),
      ),
    );
    expect([...saved.matchAll(/leftChars=/gu)]).toHaveLength(1);
    expect(saved).toContain('w:leftChars="100"');
  });
});
