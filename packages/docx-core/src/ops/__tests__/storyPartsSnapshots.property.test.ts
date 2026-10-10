import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { BlockContent, HeaderFooter } from "../../model/content";
import type { Document } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { captureStoryParts } from "../storyLifecycle";
import { DOCUMENT_OP_TYPES, type DocumentOp, type StoryParts } from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

test("generated narrow story restores preserve differing mounted snapshots and reuse equivalent parts", () => {
  assertProperty(
    fc.property(
      fc.constantFrom("default" as const, "first" as const, "even" as const),
      fc.string({ minLength: 1, maxLength: 20 }),
      (variant, text) => {
        for (const kind of ["header", "footer"] as const) {
          for (const equivalence of ["equal", "content", "source"] as const) {
            const field = kind === "header" ? "headers" : "footers";
            const references = [{ type: variant, rId: "rIdStory" }];
            const properties =
              kind === "header"
                ? { headerReferences: references }
                : { footerReferences: references };
            const owned = {
              type: kind,
              hdrFtrType: variant,
              content: [
                {
                  type: "paragraph",
                  paraId: "00000002",
                  content: [{ type: "run", content: [{ type: "text", text }] }],
                },
              ],
              verbatimXml: `<opaque>${text}</opaque>`,
              verbatimFingerprint: "owned-source",
            } satisfies HeaderFooter;
            const mounted = { ...owned, verbatimFingerprint: "mounted-source" };
            const captured = structuredClone(owned);
            if (equivalence === "content") {
              captured.content = [
                {
                  type: "paragraph",
                  paraId: "00000002",
                  content: [{ type: "run", content: [{ type: "text", text: `Captured ${text}` }] }],
                },
              ];
            }
            if (equivalence === "source") captured.verbatimFingerprint = "captured-source";
            const bodyContent = [
              { type: "paragraph", paraId: "00000001", content: [] },
            ] satisfies BlockContent[];
            const original: Document = {
              package: {
                [field]: new Map([["rIdStory", owned]]),
                document: {
                  content: [...bodyContent],
                  finalSectionProperties: properties,
                  sections: [
                    {
                      properties,
                      content: [...bodyContent],
                      [field]: new Map([[variant, mounted]]),
                    },
                  ],
                },
              },
            };
            const parts = (
              kind === "header"
                ? { sections: [{ index: 0, headers: [[variant, captured]] }] }
                : { sections: [{ index: 0, footers: [[variant, captured]] }] }
            ) satisfies StoryParts;
            const op = {
              type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
              expected: captureStoryParts(original, parts),
              parts,
            } satisfies DocumentOp;
            const expected = structuredClone(original);
            expected.package.document.sections = [
              { properties, content: [...bodyContent], [field]: new Map([[variant, captured]]) },
            ];
            const applied = applyDocumentOp(original, op).unwrap();
            expect(applied.document).toStrictEqual(expected);
            expect(applied.document.package[field]).toBe(original.package[field]);
            const restored = applied.document.package.document.sections
              ?.at(0)
              ?.[field]?.get(variant);
            const authoritative = applied.document.package[field]?.get("rIdStory");
            if (equivalence === "equal") expect(restored).toBe(authoritative);
            else expect(restored).not.toBe(authoritative);
            const undone = applyDocumentOps(applied.document, applied.inverse).unwrap();
            expect(undone.document).toStrictEqual(original);
            const redone = applyDocumentOps(undone.document, undone.inverse).unwrap();
            expect(redone.document).toStrictEqual(applied.document);
          }
        }
      },
    ),
    {
      seed: 20261019,
      numRuns: 40,
      id: "generated narrow story restores preserve differing mounted snapshots and reuse equivalent parts",
    },
  );
});
