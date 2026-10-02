/**
 * Text box enrichment preserves source order across runs lifted into comment
 * references. Existing fixed-point generators varied empty runs but omitted
 * reference runs beside drawing carriers, so their source cursor could drift.
 */
import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

const textBoxDrawing = (id: number): string =>
  '<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
  `<wp:docPr id="${id}" name="Text Box ${id}"/><a:graphic>` +
  `<a:graphicData uri="${WPS_NAMESPACE}"><wps:wsp><wps:spPr>` +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${id}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
  "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>";

const RUN_KINDS = ["reference", "textBox", "text", "textReference", "textBoxReference"] as const;
type RunKind = (typeof RUN_KINDS)[number];

const buildDocx = async (kinds: readonly RunKind[]) => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const runs = kinds.map((kind, index) => {
    switch (kind) {
      case "reference":
        return `<w:r><w:commentReference w:id="${index}"/></w:r>`;
      case "textBox":
        return `<w:r>${textBoxDrawing(index)}</w:r>`;
      case "text":
        return `<w:r><w:t>${index}</w:t></w:r>`;
      case "textReference":
        return `<w:r><w:t>${index}</w:t><w:commentReference w:id="${index}"/></w:r>`;
      case "textBoxReference":
        return `<w:r>${textBoxDrawing(index)}<w:commentReference w:id="${index}"/></w:r>`;
      default: {
        const unreachable: never = kind;
        return unreachable;
      }
    }
  });
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W_NAMESPACE}" xmlns:wp="${WP_NAMESPACE}" xmlns:a="${A_NAMESPACE}" xmlns:wps="${WPS_NAMESPACE}"><w:body><w:p>${runs.join("")}</w:p><w:sectPr/></w:body></w:document>`,
  );
  zip.file(
    "word/comments.xml",
    `<w:comments xmlns:w="${W_NAMESPACE}">${kinds.flatMap((kind, index) => (kind === "reference" || kind === "textReference" || kind === "textBoxReference" ? [`<w:comment w:id="${index}" w:author="Reviewer"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment>`] : [])).join("")}</w:comments>`,
  );
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (rels === undefined) throw new Error("Empty package lacks relationships");
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      '<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
    ),
  );
  const types = await zip.file("[Content_Types].xml")?.async("text");
  if (types === undefined) throw new Error("Empty package lacks content types");
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const payloadOrder = (document: Document) =>
  document.package.document.content.flatMap((block) => {
    if (block.type !== "paragraph") return [];
    return block.content.flatMap((item) => {
      if (item.type === "commentReference") return [`reference:${item.id}`];
      if (item.type !== "run") return [];
      return item.content.flatMap((payload) => {
        if (payload.type === "shape") {
          const text = payload.shape.textBody?.content
            .flatMap((nested) =>
              nested.type === "paragraph"
                ? nested.content.flatMap((child) =>
                    child.type === "run"
                      ? child.content.flatMap((leaf) => (leaf.type === "text" ? [leaf.text] : []))
                      : [],
                  )
                : [],
            )
            .join("");
          return [`textBox:${text}`];
        }
        if (payload.type === "text")
          return [...payload.text].map((character) => `text:${character}`);
        return [];
      });
    });
  });

test(
  "comment reference and text box run sequences preserve source order and save fixed point",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.constantFrom(...RUN_KINDS), { minLength: 1, maxLength: 8 }),
        async (kinds) => {
          const document = await parseDocx(await buildDocx(kinds));
          expect(payloadOrder(document)).toEqual(
            kinds.flatMap((kind, index) => {
              if (kind === "textReference") return [`text:${index}`, `reference:${index}`];
              if (kind === "textBoxReference") return [`textBox:${index}`, `reference:${index}`];
              return [`${kind}:${index}`];
            }),
          );
          const first = await repackDocx(document);
          const reopened = await parseDocx(first);
          expect(payloadOrder(reopened)).toEqual(payloadOrder(document));
          const second = await repackDocx(reopened);
          const firstZip = await JSZip.loadAsync(first);
          const secondZip = await JSZip.loadAsync(second);
          expect(await secondZip.file("word/document.xml")?.async("text")).toBe(
            await firstZip.file("word/document.xml")?.async("text"),
          );
        },
      ),
      { numRuns: 40 },
    );
  },
  propertyTestTimeout(30_000),
);
