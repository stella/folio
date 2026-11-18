import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { BlockContent, HeaderFooter } from "../../model/content";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { parseDocumentBatch } from "../sequencing/envelope";
import { captureStoryParts } from "../storyLifecycle";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type StoryParts,
} from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

test("resolving deleted note references restores every removed section through JSON inverses", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 1, max: 4 }),
      fc.constantFrom("footnote" as const, "endnote" as const),
      (breakCount, kind) => {
        const stamp = {
          id: 100,
          author: "Reviewer",
          date: "2026-11-16T15:00:00Z",
        };
        const boundaries = Array.from({ length: breakCount }, (_, index) => ({
          type: "paragraph",
          paraId: (index + 1).toString(16).padStart(8, "0"),
          content: [],
          sectionProperties: { pageWidth: 10000 + index },
          ...(index === breakCount - 1
            ? { pPrMark: { kind: "del", info: { ...stamp, id: 101 } } }
            : {}),
        })) satisfies Paragraph[];
        const reference: Paragraph = {
          type: "paragraph",
          paraId: "00000010",
          content: [
            {
              type: "deletion",
              info: stamp,
              content: [
                {
                  type: "run",
                  content: [
                    {
                      type: kind === "footnote" ? "footnoteRef" : "endnoteRef",
                      id: 2,
                    },
                  ],
                },
              ],
              resolutionJoins: { before: 0, after: 0, remove: 0 },
            },
          ],
        };
        const properties = { pageWidth: 12000 };
        const noteContent = [
          { type: "paragraph", paraId: "00000011", content: [] },
        ] satisfies Paragraph[];
        const document: Document = {
          package: {
            document: {
              content: [...boundaries, reference],
              finalSectionProperties: properties,
              sections: [
                ...boundaries.map((paragraph) => ({
                  properties: paragraph.sectionProperties,
                  content: [paragraph],
                })),
                { properties, content: [reference] },
              ],
            },
          },
        };
        if (kind === "footnote")
          document.package.footnotes = [
            { type: kind, id: 2, content: noteContent },
          ];
        else
          document.package.endnotes = [
            { type: kind, id: 2, content: noteContent },
          ];
        const resolved = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [101, 100],
          decision: REVISION_DECISIONS.ACCEPT,
        }).unwrap();
        expect(resolved.document.package.document.sections).toHaveLength(
          breakCount,
        );
        expect(
          resolved.document.package[
            kind === "footnote" ? "footnotes" : "endnotes"
          ],
        ).toEqual([]);
        const decoded = parseDocumentBatch(
          JSON.stringify({
            schema: DOCUMENT_OP_SCHEMA_VERSION,
            opId: "section-lifecycle-undo",
            actor: "reviewer",
            baseRev: 0,
            ops: resolved.inverse,
          }),
        ).unwrap();
        const undone = applyDocumentOps(
          resolved.document,
          decoded.ops,
        ).unwrap();
        expect(undone.document).toStrictEqual(document);
        const redone = applyDocumentOps(
          undone.document,
          JSON.parse(JSON.stringify(undone.inverse)),
        ).unwrap();
        expect(redone.document).toStrictEqual(resolved.document);
        const later = applyDocumentOp(resolved.document, {
          type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
          sectionIndex: breakCount - 1,
          patch: { marginLeft: 777 },
        }).unwrap();
        const stale = applyDocumentOps(later.document, decoded.ops);
        expect(stale.isErr()).toBe(true);
        if (stale.isErr())
          expect(stale.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
        const inverse = decoded.ops.at(0);
        expect(inverse?.type).toBe(DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS);
        if (inverse?.type !== DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS) return;
        const malformed = applyDocumentOp(resolved.document, {
          type: inverse.type,
          expected: inverse.expected,
          parts: {
            body: inverse.parts.body,
            sectionView: { type: "sections", value: [] },
          },
        });
        expect(malformed.isErr()).toBe(true);
        if (malformed.isErr())
          expect(malformed.error.reason).toBe(
            DOCUMENT_OP_REFUSAL_REASONS.STALE,
          );
      },
    ),
    {
      seed: 20261116,
      numRuns: 16,
      id: "note resolution restores section topology through JSON inverses",
    },
  );
});

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
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: `Captured ${text}` }],
                    },
                  ],
                },
              ];
            }
            if (equivalence === "source")
              captured.verbatimFingerprint = "captured-source";
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
              {
                properties,
                content: [...bodyContent],
                [field]: new Map([[variant, captured]]),
              },
            ];
            const applied = applyDocumentOp(original, op).unwrap();
            expect(applied.document).toStrictEqual(expected);
            expect(applied.document.package[field]).toBe(
              original.package[field],
            );
            const restored = applied.document.package.document.sections
              ?.at(0)
              ?.[field]?.get(variant);
            const authoritative =
              applied.document.package[field]?.get("rIdStory");
            if (equivalence === "equal") expect(restored).toBe(authoritative);
            else expect(restored).not.toBe(authoritative);
            const undone = applyDocumentOps(
              applied.document,
              applied.inverse,
            ).unwrap();
            expect(undone.document).toStrictEqual(original);
            const redone = applyDocumentOps(
              undone.document,
              undone.inverse,
            ).unwrap();
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
