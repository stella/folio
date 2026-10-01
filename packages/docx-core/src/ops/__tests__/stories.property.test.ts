/** Story generators extend the operation oracle beyond the main body. */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { propertyConfig } from "../../../../../test/property-testing";
import type { Document, Paragraph, HeaderFooter, SectionProperties } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation, normalizeForOps } from "../contract";
import { documentStories, findStoryBody, sameStory } from "../stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type OpStory,
  INHERIT_RUN_PROPS,
  SPLIT_HALVES,
} from "../types";

import { documentArbitrary, opSeedArbitrary, opForStory } from "./documentArbitraries";

const paragraph = (id: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId: id,
  content: [{ type: "run", formatting: { bold: true }, content: [{ type: "text", text }] }],
});
const seedDocument = (text: string): Document => ({
  package: {
    document: { content: [paragraph("00000001", text)] },
    headers: new Map([
      [
        "rIdHeader",
        { type: "header", hdrFtrType: "default", content: [paragraph("00000002", text)] },
      ],
    ]),
    footers: new Map([
      [
        "rIdFooter",
        { type: "footer", hdrFtrType: "default", content: [paragraph("00000003", text)] },
      ],
    ]),
    footnotes: [{ type: "footnote", id: 1, content: [paragraph("00000004", text)] }],
    endnotes: [{ type: "endnote", id: 1, content: [paragraph("00000005", text)] }],
  },
});

const expectExact = (document: Document, op: DocumentOp) => {
  const applied = applyDocumentOp(document, op).unwrap();
  expect(contractViolation(applied.document)).toBeUndefined();
  const wireOp: DocumentOp = JSON.parse(JSON.stringify(op));
  expect(applyDocumentOp(structuredClone(document), wireOp).unwrap()).toEqual(applied);
  const inverse: DocumentOp[] = JSON.parse(JSON.stringify(applied.inverse));
  const undone = applyDocumentOps(applied.document, inverse).unwrap();
  expect(undone.document).toEqual(document);
  expect(applyDocumentOps(undone.document, undone.inverse).unwrap().document).toEqual(
    applied.document,
  );
  return applied;
};

describe("all-story operation laws", () => {
  test("generated text, formatting, split and join operations are exact and local in every story", () => {
    fc.assert(
      fc.property(fc.stringMatching(/^[A-Za-z]{2,12}$/), fc.nat(), (text, selector) => {
        const document = normalizeForOps(seedDocument(text));
        for (const story of documentStories(document)) {
          const body = findStoryBody(document, story);
          const block = body?.content.at(0);
          if (block?.type !== "paragraph" || !block.paraId)
            return panic("Generated story must hold an identified paragraph.");
          const at = { story, blockId: block.paraId, offset: selector % text.length };
          const ops = [
            { type: DOCUMENT_OP_TYPES.INSERT_TEXT, at, text: "Z", runProps: INHERIT_RUN_PROPS },
            {
              type: DOCUMENT_OP_TYPES.DELETE_RANGE,
              from: at,
              to: { ...at, story: structuredClone(story), offset: at.offset + 1 },
            },
            {
              type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
              from: at,
              to: { ...at, offset: at.offset + 1 },
              patch: { italic: true },
            },
            {
              type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
              story,
              blockId: block.paraId,
              patch: { alignment: "center" },
            },
            {
              type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
              at,
              newBlockId: "000000A0",
              newHalf: SPLIT_HALVES.SECOND,
            },
          ] as const satisfies readonly DocumentOp[];
          for (const op of ops) {
            const applied = expectExact(document, op);
            for (const untouched of documentStories(document)) {
              if (sameStory(story, untouched)) continue;
              expect(findStoryBody(applied.document, untouched)).toBe(
                findStoryBody(document, untouched),
              );
            }
            if (op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK) {
              expectExact(applied.document, {
                type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
                story: structuredClone(story),
                blockId: block.paraId,
                nextBlockId: "000000A0",
              });
            }
          }
        }
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("the existing rich operation generator exercises every other story with the same inverse and locality oracle", () => {
    let appliedCount = 0;
    fc.assert(
      fc.property(documentArbitrary, opSeedArbitrary, (source, seed) => {
        const stories = [
          { kind: "header", rId: "rIdGenerated" },
          { kind: "footer", rId: "rIdGenerated" },
          { kind: "footnote", id: 99 },
          { kind: "endnote", id: 99 },
        ] as const satisfies readonly OpStory[];
        for (const story of stories) {
          const content = source.package.document.content;
          const main = { content: [paragraph("7FFFFFFE", "unchanged")] };
          const pkg = { ...source.package, document: main };
          const document = (() => {
            switch (story.kind) {
              case "header":
                return {
                  ...source,
                  package: {
                    ...pkg,
                    headers: new Map([
                      [story.rId, { type: "header", hdrFtrType: "default", content }],
                    ]),
                  },
                } satisfies Document;
              case "footer":
                return {
                  ...source,
                  package: {
                    ...pkg,
                    footers: new Map([
                      [story.rId, { type: "footer", hdrFtrType: "default", content }],
                    ]),
                  },
                } satisfies Document;
              case "footnote":
                return {
                  ...source,
                  package: {
                    ...pkg,
                    footnotes: [
                      ...(pkg.footnotes ?? []),
                      { type: "footnote", id: story.id, content },
                    ],
                  },
                } satisfies Document;
              case "endnote":
                return {
                  ...source,
                  package: { ...pkg, endnotes: [{ type: "endnote", id: story.id, content }] },
                } satisfies Document;
              default: {
                const unreachable: never = story;
                return unreachable;
              }
            }
          })();
          const op = opForStory({ document, seed, story });
          const result = applyDocumentOp(document, op);
          expect(result).toEqual(
            applyDocumentOp(structuredClone(document), JSON.parse(JSON.stringify(op))),
          );
          if (result.isErr()) continue;
          appliedCount += 1;
          const applied = expectExact(document, op);
          for (const untouched of documentStories(document)) {
            if (sameStory(story, untouched)) continue;
            expect(findStoryBody(applied.document, untouched)).toBe(
              findStoryBody(document, untouched),
            );
          }
        }
      }),
      propertyConfig({ numRuns: 100 }),
    );
    expect(appliedCount).toBeGreaterThan(50);
  });

  test("generated header/footer variants, note references and section properties have exact inverses", () => {
    fc.assert(
      fc.property(
        fc.constantFrom("header", "footer"),
        fc.constantFrom("default", "first", "even"),
        fc.constantFrom("footnote", "endnote"),
        fc.integer({ min: 1, max: 100 }),
        (kind, referenceType, noteKind, id) => {
          const document: Document = {
            package: { document: { content: [paragraph("00000001", "body")] } },
          };
          const created = expectExact(document, {
            type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
            sectionIndex: 0,
            story: { kind, rId: "rIdNew" },
            referenceType,
            content: [paragraph("00000002", "story")],
          });
          expectExact(created.document, {
            type: DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
            sectionIndex: 0,
            story: { kind, rId: "rIdNew" },
            referenceType,
          });
          if (referenceType === "first")
            expect(created.document.package.document.finalSectionProperties?.titlePg).toBe(true);
          if (referenceType === "even")
            expect(created.document.package.settings?.evenAndOddHeaders).toBe(true);
          const at = { story: "main", blockId: "00000001", offset: 1 } as const;
          const added = expectExact(document, {
            type: DOCUMENT_OP_TYPES.ADD_NOTE,
            at,
            note: { type: noteKind, id, content: [paragraph("00000003", "note")] },
          });
          expectExact(added.document, {
            type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
            at,
            story: { kind: noteKind, id },
          });
          expectExact(document, {
            type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
            sectionIndex: 0,
            patch: { footnotePr: { numStart: id }, marginTop: id * 20 },
          });
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("shared header parts survive variant removal and section mirrors follow story edits after wire inverses", () => {
    const block = paragraph("00000001", "body");
    const header = {
      type: "header",
      hdrFtrType: "default",
      content: [paragraph("00000002", "header")],
    } as const satisfies HeaderFooter;
    const properties = {
      headerReferences: [
        { type: "default", rId: "rIdShared" },
        { type: "first", rId: "rIdShared" },
      ],
    } as const satisfies SectionProperties;
    const document: Document = {
      package: {
        document: {
          content: [block],
          finalSectionProperties: properties,
          sections: [
            {
              properties,
              content: [block],
              headers: new Map([
                ["default", header],
                ["first", header],
              ]),
            },
          ],
        },
        headers: new Map([["rIdShared", header]]),
      },
    };
    const removed = expectExact(document, {
      type: DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
      sectionIndex: 0,
      story: { kind: "header", rId: "rIdShared" },
      referenceType: "default",
    });
    expect(removed.document.package.headers?.has("rIdShared")).toBe(true);
    expectExact(removed.document, {
      type: DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
      sectionIndex: 0,
      story: { kind: "header", rId: "rIdShared" },
      referenceType: "first",
    });
    const wireInverse: DocumentOp[] = JSON.parse(JSON.stringify(removed.inverse));
    const restored = applyDocumentOps(removed.document, wireInverse).unwrap().document;
    const edited = expectExact(restored, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story: { kind: "header", rId: "rIdShared" }, blockId: "00000002", offset: 0 },
      text: "Z",
      runProps: INHERIT_RUN_PROPS,
    });
    expect(edited.document.package.document.sections?.at(0)?.headers?.get("first")).toBe(
      edited.document.package.headers?.get("rIdShared"),
    );
  });

  test("new notes own one zero-width automatic mark before serialization", () => {
    for (const kind of ["footnote", "endnote"] as const) {
      const document: Document = {
        package: { document: { content: [paragraph("00000001", "body")] } },
      };
      const added = expectExact(document, {
        type: DOCUMENT_OP_TYPES.ADD_NOTE,
        at: { story: "main", blockId: "00000001", offset: 0 },
        note: { type: kind, id: 1, content: [paragraph("00000002", "note")] },
      });
      const note = findStoryBody(added.document, { kind, id: 1 })?.content.at(0);
      if (note?.type !== "paragraph") return panic("The new note has a paragraph.");
      expect(note.content.at(0)).toEqual({
        type: "run",
        formatting: { styleId: kind === "footnote" ? "FootnoteReference" : "EndnoteReference" },
        content: [{ type: "noteMarker", kind }],
      });
      const edited = expectExact(added.document, {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: { story: { kind, id: 1 }, blockId: "00000002", offset: 0 },
        text: "Z",
        runProps: INHERIT_RUN_PROPS,
      });
      const after = findStoryBody(edited.document, { kind, id: 1 })?.content.at(0);
      if (after?.type !== "paragraph") return panic("The edited note has a paragraph.");
      expect(after.content.at(0)).toEqual(note.content.at(0));
    }
  });

  test("normalization derives automatic note marks and preserves custom reference marks and separators", () => {
    for (const kind of ["footnote", "endnote"] as const) {
      const referenceType = kind === "footnote" ? "footnoteRef" : "endnoteRef";
      for (const customMarkFollows of [false, true]) {
        const document: Document = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [
                    { type: "run", content: [{ type: referenceType, id: 1, customMarkFollows }] },
                  ],
                },
              ],
            },
            ...(kind === "footnote"
              ? {
                  footnotes: [
                    { type: "footnote", id: 1, content: [paragraph("00000002", "note")] },
                    {
                      type: "footnote",
                      id: -1,
                      noteType: "separator",
                      content: [paragraph("00000003", "separator")],
                    },
                  ],
                }
              : {
                  endnotes: [
                    { type: "endnote", id: 1, content: [paragraph("00000002", "note")] },
                    {
                      type: "endnote",
                      id: -1,
                      noteType: "separator",
                      content: [paragraph("00000003", "separator")],
                    },
                  ],
                }),
          },
        };
        const normalized = normalizeForOps(document);
        const note = findStoryBody(normalized, { kind, id: 1 })?.content.at(0);
        if (note?.type !== "paragraph") return panic("The normalized note must exist.");
        expect(note.content.at(0)?.type).toBe("run");
        const first = note.content.at(0);
        const mark = first?.type === "run" ? first.content.at(0) : undefined;
        expect(mark?.type).toBe(customMarkFollows ? "text" : "noteMarker");
        expect(findStoryBody(normalized, { kind, id: -1 })).toEqual(
          findStoryBody(document, { kind, id: -1 }),
        );
        expect(normalizeForOps(normalized)).toEqual(normalized);
      }
    }
  });

  test("normalization and validation cover every editable story", () => {
    const document = seedDocument("body");
    const header = document.package.headers?.get("rIdHeader");
    if (!header) return panic("Generated header must exist.");
    const first = header.content.at(0);
    if (first?.type !== "paragraph") return panic("Generated paragraph must exist.");
    first.content.push({ type: "run", content: [{ type: "text", text: "" }] });
    expect(contractViolation(document)?.reason).toBe("emptyRecord");
    expect(contractViolation(normalizeForOps(document))).toBeUndefined();
    delete first.paraId;
    expect(contractViolation(document)?.reason).toBe("missingBlockId");
  });

  test("lifecycle inverse refuses changed section or story contents without losing intervening edits", () => {
    const document = seedDocument("body");
    const changed = expectExact(document, {
      type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
      sectionIndex: 0,
      patch: { marginTop: 720 },
    });
    const intervening = applyDocumentOp(changed.document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story: { kind: "header", rId: "rIdHeader" }, blockId: "00000002", offset: 0 },
      text: "Z",
      runProps: INHERIT_RUN_PROPS,
    }).unwrap();
    const stale = applyDocumentOps(intervening.document, changed.inverse);
    expect(stale.isErr()).toBe(true);
    expect(
      findStoryBody(intervening.document, { kind: "header", rId: "rIdHeader" })?.content.at(0),
    ).toEqual(paragraph("00000002", "Zbody"));
  });

  test("missing stories and stale note references refuse atomically", () => {
    const document = normalizeForOps(seedDocument("body"));
    const story: OpStory = { kind: "header", rId: "missing" };
    expect(
      applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: { story, blockId: "00000001", offset: 0 },
        text: "X",
        runProps: INHERIT_RUN_PROPS,
      }).isErr(),
    ).toBe(true);
    expect(
      applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
        at: { story: "main", blockId: "00000001", offset: 0 },
        story: { kind: "footnote", id: 1 },
      }).isErr(),
    ).toBe(true);
    expect(findStoryBody(document, "main")?.content.at(0)).toEqual(paragraph("00000001", "body"));
  });
});
