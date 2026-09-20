/**
 * A field character that belongs to no field must survive, and so must the run
 * around it.
 *
 * `w:fldChar` is half of a construct: a `begin`, an instruction, a `separate`
 * and an `end` spread over several runs assemble into one `ComplexField`. The
 * paragraph parser's state machine buffered every run it read after a `begin`
 * and handed the buffer to the field it built at the `end` — so a paragraph
 * that ended with the field still open threw the buffer away. The census
 * measured it as `the-container-itself-is-lost` for all 13 `w:ffData` pairs,
 * but the visible cost is larger: the runs it dropped carried their `w:t` too,
 * so opening and saving a document with a `FORMCHECKBOX` field whose `end`
 * Word wrote in the next paragraph deleted the text beside it.
 *
 * What the character carries is the other half. `CT_FFData` holds a legacy form
 * field's name, its enable and recalculate flags, its entry and exit macros,
 * its help and status text and the checkbox, dropdown or text-input state, and
 * the model holds none of it: `FieldCharContent` is a `charType` and two
 * `ST_OnOff` flags. The run's verbatim capture member carries the element
 * whole, which is why this property asks about every child the schema declares
 * rather than the handful a reader could name.
 *
 * Both legs matter and they fail differently. The save leg was losing the run;
 * the editor leg drops a bare `FieldCharContent` in the inline converter, so a
 * fix that only reached the serializer would be undone by the first round trip.
 */

import { readFileSync } from "node:fs";

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type SchemaGraph = {
  children: Array<{ owner: string; kind?: string; name?: string; ref?: string }>;
  symbols: Array<{ id: string; kind?: string; name?: string }>;
};

/**
 * `CT_FFData`'s declared children, read from the committed schema graph rather
 * than restated: a schema refresh that adds one widens this property instead of
 * leaving a hand list behind to drift.
 */
const FF_DATA_CHILDREN: readonly string[] = (() => {
  const graph = JSON.parse(
    readFileSync(
      new URL(
        "../../../../specifications/generated/docx-transitional-schema.gen.json",
        import.meta.url,
      ),
      "utf8",
    ),
  ) as SchemaGraph;
  const byId = new Map(graph.symbols.map((symbol) => [symbol.id, symbol]));
  const owner = `complexType:{${W_NAMESPACE}}CT_FFData`;
  const declared = graph.children
    .filter((child) => child.owner === owner && child.kind === "element")
    .map((child) => (child.ref === undefined ? child.name : byId.get(`element:${child.ref}`)?.name))
    .filter((name): name is string => name !== undefined);
  if (declared.length === 0) {
    throw new Error("CT_FFData declares no children in the schema graph");
  }
  return declared;
})();

/**
 * One child of `w:ffData`, as a document writes it.
 *
 * The capture is verbatim, so the markup only has to be what Word would accept;
 * what is under test is whether it comes back at all. The totality check below
 * is the guard: a child the schema declares and this table does not name fails
 * the suite rather than silently narrowing the property.
 */
const FF_DATA_CHILD_MARKUP: Readonly<Record<string, string>> = {
  name: '<w:name w:val="folioField"/>',
  label: '<w:label w:val="7"/>',
  tabIndex: '<w:tabIndex w:val="3"/>',
  enabled: '<w:enabled w:val="0"/>',
  calcOnExit: '<w:calcOnExit w:val="1"/>',
  entryMacro: '<w:entryMacro w:val="OnEntry"/>',
  exitMacro: '<w:exitMacro w:val="OnExit"/>',
  helpText: '<w:helpText w:type="text" w:val="what to type"/>',
  statusText: '<w:statusText w:type="text" w:val="on the status bar"/>',
  checkBox: '<w:checkBox><w:size w:val="24"/><w:default w:val="1"/></w:checkBox>',
  ddList: '<w:ddList><w:result w:val="1"/><w:listEntry w:val="one"/></w:ddList>',
  textInput: '<w:textInput><w:type w:val="regular"/><w:default w:val="typed"/></w:textInput>',
};

const FIELD_CHAR_TYPES = ["begin", "separate", "end"] as const;

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentXmlOf = async (saved: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";

/** The part a plain save writes, and the part the editor round trip writes. */
const savedAndProjected = async (body: string): Promise<{ saved: string; projected: string }> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  const saved = await documentXmlOf(await repackDocx(parsed, { updateModifiedDate: false }));
  const roundTripped = fromProseDoc(toProseDoc(parsed), parsed);
  const projected = await documentXmlOf(
    await repackDocx(roundTripped, { updateModifiedDate: false }),
  );
  return { saved, projected };
};

const fieldCharRun = (charType: string, inner: string): string =>
  `<w:r><w:t>beside</w:t><w:fldChar w:fldCharType="${charType}">${inner}</w:fldChar></w:r>`;

describe("a field character with no field keeps itself and its run", () => {
  test("every child w:ffData declares has markup here", () => {
    expect([...FF_DATA_CHILDREN].sort()).toEqual(
      Object.keys(FF_DATA_CHILD_MARKUP).sort((left, right) => left.localeCompare(right)),
    );
  });

  test("every w:ffData child survives a save and the editor", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...FF_DATA_CHILDREN),
        fc.constantFrom(...FIELD_CHAR_TYPES),
        async (childName, charType) => {
          const markup = FF_DATA_CHILD_MARKUP[childName];
          expect(markup).toBeDefined();
          const body = `<w:p>${fieldCharRun(charType, `<w:ffData>${markup ?? ""}</w:ffData>`)}</w:p>`;
          const { saved, projected } = await savedAndProjected(body);

          // The child, the element that holds it, and the text the dropped run
          // used to take with it.
          for (const part of [saved, projected]) {
            expect({ childName, charType, part: part.includes(markup ?? "") }).toEqual({
              childName,
              charType,
              part: true,
            });
            expect(part).toContain(`w:fldCharType="${charType}"`);
            expect(part).toContain("<w:t>beside</w:t>");
          }
        },
      ),
      propertyConfig({ numRuns: 36 }),
    );
  });

  test("a field opened and never closed keeps its instruction run's text", async () => {
    const body =
      "<w:p>" +
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="cb"/>' +
      '<w:checkBox><w:default w:val="1"/></w:checkBox></w:ffData></w:fldChar></w:r>' +
      "<w:r><w:instrText> FORMCHECKBOX </w:instrText></w:r>" +
      "<w:r><w:t>label text</w:t></w:r>" +
      "</w:p>";
    const { saved, projected } = await savedAndProjected(body);

    for (const part of [saved, projected]) {
      expect(part).toContain('<w:name w:val="cb"/>');
      expect(part).toContain('<w:default w:val="1"/>');
      expect(part).toContain("<w:t>label text</w:t>");
    }
  });
});
