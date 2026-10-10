import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

test("generated lifecycle inverses preserve absent and own-undefined maps and final properties", () => {
  assertProperty(
    fc.property(
      fc.constantFrom("headers" as const, "footers" as const),
      fc.constantFrom("story" as const, "section" as const),
      fc.boolean(),
      (field, kind, explicit) => {
        const original: Document = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [{ type: "run", content: [{ type: "text", text: "Body" }] }],
                },
              ],
            },
          },
        };
        if (explicit) {
          Reflect.set(original.package, field, undefined);
          Reflect.set(original.package.document, "finalSectionProperties", undefined);
        }
        const op: DocumentOp =
          kind === "section"
            ? {
                type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
                sectionIndex: 0,
                patch: { footnotePr: { numStart: 3 }, marginLeft: 800 },
              }
            : {
                type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
                sectionIndex: 0,
                referenceType: "default",
                story: { kind: field === "headers" ? "header" : "footer", rId: "rIdFresh" },
                content: [{ type: "paragraph", paraId: "00000002", content: [] }],
              };
        const applied = applyDocumentOp(original, op).unwrap();
        const undone = applyDocumentOps(
          applied.document,
          JSON.parse(JSON.stringify(applied.inverse)),
        ).unwrap();
        expect(undone.document).toStrictEqual(original);
        expect(Object.hasOwn(undone.document.package, field)).toBe(explicit);
        expect(Object.hasOwn(undone.document.package.document, "finalSectionProperties")).toBe(
          explicit,
        );
        const redone = applyDocumentOps(
          undone.document,
          JSON.parse(JSON.stringify(undone.inverse)),
        ).unwrap();
        expect(redone.document).toStrictEqual(applied.document);
      },
    ),
    {
      seed: 20261015,
      numRuns: 40,
      id: "generated lifecycle inverses preserve absent and own-undefined maps and final properties",
    },
  );
});
