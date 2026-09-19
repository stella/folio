/**
 * What a saved document says must be what the source said.
 *
 * Every other fidelity invariant is about formatting or structure; this one is
 * about the words. The public corpus found two ways folio broke it, and both
 * are shapes a generator reaches: a `w:pict` represented twice wrote its text
 * twice, and a result-less `PAGE` field gained a literal page number the author
 * never wrote.
 *
 * The property runs both save paths, because they lose differently: a plain
 * repack rebuilds every paragraph from the model, and the editor round trip
 * rebuilds it from ProseMirror. The expected text comes from the generator
 * rather than from a second folio walk, so a loss on the read side cannot
 * cancel out with a loss on the write side.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, Document } from "../types/document";
import { getParagraphText } from "./paragraphParser";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** A run child, the visible text it contributes, and what it is called. */
type RunChild = { xml: string; text: string };

const textChild = (raw: string, preserve: boolean): RunChild => ({
  xml: `<w:t${preserve ? ' xml:space="preserve"' : ""}>${raw}</w:t>`,
  text: raw,
});

/**
 * `w:t` payloads whose meaning depends on `xml:space`: a reader trims a leading
 * or trailing space and collapses an inner run of them unless the attribute
 * says otherwise, so folio has to write the attribute back whenever it matters.
 */
const TEXT_PAYLOADS = ["plain", " leading", "trailing ", "  ", "", "two  spaces"] as const;

const ATOM_CHILDREN = {
  tab: { xml: "<w:tab/>", text: "\t" },
  lineBreak: { xml: "<w:br/>", text: "\n" },
  pageBreak: { xml: '<w:br w:type="page"/>', text: "\f" },
  softHyphen: { xml: "<w:softHyphen/>", text: "­" },
  noBreakHyphen: { xml: "<w:noBreakHyphen/>", text: "‑" },
  symbol: { xml: '<w:sym w:font="Wingdings" w:char="F0E0"/>', text: "" },
  renderedPageBreak: { xml: "<w:lastRenderedPageBreak/>", text: "" },
} as const satisfies Record<string, RunChild>;

/**
 * A VML group holding `count` text boxes.
 *
 * The run parser preserves the whole `w:pict` as one raw drawing (its group
 * renders a preview), so the text-box pass must not also rebuild the first
 * box as an editable shape: that second representation wrote the box's text a
 * second time on every save.
 */
const vmlGroupPict = (count: number): RunChild => {
  const boxes = Array.from(
    { length: count },
    (_, index) =>
      `<v:shape id="s${String(index)}" style="position:absolute;left:0;top:${String(
        index * 100,
      )};width:100pt;height:20pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>box${String(
        index,
      )}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>`,
  ).join("");
  return {
    // The oval is what the group preview paints, in the group's own
    // coordinate space; without a painted child the preview declines and the
    // pict never reaches the run parser's claim.
    xml: `<w:pict><v:group id="g" style="position:absolute;width:200pt;height:100pt" coordorigin="0,0" coordsize="2000,1000"><v:oval id="o" style="position:absolute;left:0;top:0;width:500;height:500"/>${boxes}</v:group></w:pict>`,
    // A text box is page artwork, not paragraph text: neither the source nor
    // the saved package contributes it to the paragraph's own text.
    text: "",
  };
};

/** A bare VML text box: no group preview claims it, so the model holds it. */
const vmlShapePict: RunChild = {
  xml: '<w:pict><v:shape id="bare" style="position:absolute;width:100pt;height:20pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>bare box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>',
  text: "bare box",
};

/** A complex `PAGE` field with an empty result: Word computes the number. */
const RESULT_LESS_PAGE_FIELD: RunChild = {
  xml: '<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/>',
  text: "",
};

type ChildSpec =
  | { kind: "text"; payload: (typeof TEXT_PAYLOADS)[number]; preserve: boolean }
  | { kind: "atom"; name: keyof typeof ATOM_CHILDREN }
  | { kind: "vmlGroup"; boxes: number }
  | { kind: "vmlShape" }
  | { kind: "pageField" };

const childFor = (spec: ChildSpec): RunChild => {
  switch (spec.kind) {
    case "text":
      return textChild(spec.payload, spec.preserve);
    case "atom":
      return ATOM_CHILDREN[spec.name];
    case "vmlGroup":
      return vmlGroupPict(spec.boxes);
    case "vmlShape":
      return vmlShapePict;
    case "pageField":
      return RESULT_LESS_PAGE_FIELD;
  }
};

const childArbitrary: fc.Arbitrary<ChildSpec> = fc.oneof(
  fc.record({
    kind: fc.constant<"text">("text"),
    payload: fc.constantFrom(...TEXT_PAYLOADS),
    preserve: fc.boolean(),
  }),
  fc.record({
    kind: fc.constant<"atom">("atom"),
    name: fc.constantFrom(...(Object.keys(ATOM_CHILDREN) as (keyof typeof ATOM_CHILDREN)[])),
  }),
  fc.record({ kind: fc.constant<"vmlGroup">("vmlGroup"), boxes: fc.integer({ min: 1, max: 3 }) }),
  fc.record({ kind: fc.constant<"vmlShape">("vmlShape") }),
  fc.record({ kind: fc.constant<"pageField">("pageField") }),
);

const paragraphArbitrary = fc.array(childArbitrary, { minLength: 1, maxLength: 5 });

const documentXml = (children: readonly RunChild[]): string => `${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body><w:p>${children.map(({ xml }) => `<w:r>${xml}</w:r>`).join("")}</w:p><w:sectPr/></w:body>
</w:document>`;

const buildPackage = async (children: readonly RunChild[]): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", documentXml(children));
  return zip.generateAsync({ type: "arraybuffer" });
};

/**
 * Every word the model holds for these blocks, text boxes included.
 *
 * A text box's words are not the host paragraph's text, but they are the
 * document's: the defect that motivated this property wrote a text box twice,
 * and a walk that stopped at the paragraph would not have seen it.
 */
const bodyText = (blocks: readonly BlockContent[]): string => {
  const parts: string[] = [];
  const visit = (block: BlockContent): void => {
    if (block.type === "paragraph") {
      parts.push(getParagraphText(block));
      for (const content of block.content) {
        if (content.type !== "run") {
          continue;
        }
        for (const runContent of content.content) {
          if (runContent.type === "shape" && runContent.shape.textBody) {
            for (const child of runContent.shape.textBody.content) visit(child);
          }
        }
      }
      return;
    }
    if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) for (const child of cell.content) visit(child);
      }
      return;
    }
    for (const child of block.content) visit(child);
  };
  for (const block of blocks) visit(block);
  return parts.join("");
};

/**
 * The text the generator wrote. A `w:t` whose payload needs
 * `xml:space="preserve"` gets the attribute back from the serializer whether or
 * not the source carried it, so the space a reader keeps does not depend on the
 * save.
 *
 * A VML group is excluded: folio keeps it as captured XML rather than as model
 * content, so the package still says its words while the model does not read
 * them. What must hold for it is the fixed point, which is asserted separately.
 */
const statedSourceText = (children: readonly ChildSpec[]): string | undefined => {
  if (children.some((spec) => spec.kind === "vmlGroup")) {
    return undefined;
  }
  // `bodyText` reports a paragraph's own words and then the words of the boxes
  // it anchors, because a box is placed rather than laid out in the run order.
  const anchored = children.filter((spec) => spec.kind === "vmlShape");
  const inline = children.filter((spec) => spec.kind !== "vmlShape");
  return [...inline, ...anchored].map((spec) => childFor(spec).text).join("");
};

const repack = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

describe("a saved package says what the source said", () => {
  test("a plain repack preserves the paragraph's text", async () => {
    await fc.assert(
      fc.asyncProperty(paragraphArbitrary, async (specs) => {
        const parsed = await parseDocx(await buildPackage(specs.map(childFor)), {
          preloadFonts: false,
        });
        const read = bodyText(parsed.package.document.content);
        const stated = statedSourceText(specs);
        if (stated !== undefined) {
          expect(read).toBe(stated);
        }

        const saved = await parseDocx(await repack(parsed), { preloadFonts: false });
        expect(bodyText(saved.package.document.content)).toBe(read);
      }),
      propertyConfig({ numRuns: 120 }),
    );
  });

  test("the editor round trip preserves the paragraph's text", async () => {
    await fc.assert(
      fc.asyncProperty(paragraphArbitrary, async (specs) => {
        const parsed = await parseDocx(await buildPackage(specs.map(childFor)), {
          preloadFonts: false,
        });
        const read = bodyText(parsed.package.document.content);
        const stated = statedSourceText(specs);
        if (stated !== undefined) {
          expect(read).toBe(stated);
        }

        const rebuilt = fromProseDoc(toProseDoc(parsed), parsed);
        const saved = await parseDocx(await repack(rebuilt), { preloadFonts: false });
        expect(bodyText(saved.package.document.content)).toBe(read);
      }),
      propertyConfig({ numRuns: 120 }),
    );
  });
});
