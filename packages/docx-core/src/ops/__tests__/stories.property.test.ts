/** Story generators extend the operation oracle beyond the main body. */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph, HeaderFooter, SectionProperties } from "../../model/document";
import { DEFAULT_TAB_STOP_TWIPS } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { applyStoryLifecycle } from "../storyLifecycle";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
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

setDefaultTimeout(propertyTestTimeout(30_000));

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
        fc.option(
          fc.record({
            defaultTabStop: fc.integer({ min: 1, max: 2880 }),
            mirrorMargins: fc.boolean(),
            evenAndOddHeaders: fc.boolean(),
          }),
          { nil: undefined },
        ),
        (kind, referenceType, noteKind, id, settings) => {
          const document: Document = {
            package: {
              document: { content: [paragraph("00000001", "body")] },
              ...(settings === undefined ? {} : { settings }),
            },
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
            expect(created.document.package.settings).toEqual({
              ...(settings ?? { defaultTabStop: DEFAULT_TAB_STOP_TWIPS }),
              evenAndOddHeaders: true,
            });
          else expect(created.document.package.settings).toEqual(settings);
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
          for (const evenAndOddHeaders of [true, false, null]) {
            const changed = expectExact(document, {
              type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
              sectionIndex: 0,
              patch: { evenAndOddHeaders },
            });
            if (evenAndOddHeaders === null) {
              expect(changed.document.package.settings).toEqual(
                settings === undefined
                  ? undefined
                  : {
                      defaultTabStop: settings.defaultTabStop,
                      mirrorMargins: settings.mirrorMargins,
                    },
              );
            } else {
              expect(changed.document.package.settings).toEqual({
                ...(settings ?? { defaultTabStop: DEFAULT_TAB_STOP_TWIPS }),
                evenAndOddHeaders,
              });
            }
          }
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
    // Existing staleness fixtures changed values but never removed the owned section itself.
    for (const sections of [undefined, []]) {
      const body = { ...removed.document.package.document };
      if (sections === undefined) delete body.sections;
      else body.sections = sections;
      const missingSection = {
        ...removed.document,
        package: { ...removed.document.package, document: body },
      };
      for (const inverse of removed.inverse) {
        if (inverse.type !== DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS)
          return panic("Header removal must have a lifecycle inverse.");
        const result = applyStoryLifecycle(missingSection, inverse);
        expect(result.isErr()).toBe(true);
        if (result.isErr()) expect(result.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
      }
      expect(missingSection.package.document.sections).toEqual(sections);
    }
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
      type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
      sectionIndex: 0,
      patch: { marginTop: 1440 },
    }).unwrap();
    const stale = applyDocumentOps(intervening.document, changed.inverse);
    expect(stale.isErr()).toBe(true);
    expect(intervening.document.package.document.finalSectionProperties?.marginTop).toBe(1440);

    const created = expectExact(document, {
      type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
      sectionIndex: 0,
      story: { kind: "header", rId: "rIdNew" },
      referenceType: "default",
      content: [paragraph("00000006", "new")],
    });
    const edited = applyDocumentOp(created.document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story: { kind: "header", rId: "rIdNew" }, blockId: "00000006", offset: 0 },
      text: "Z",
      runProps: INHERIT_RUN_PROPS,
    }).unwrap();
    expect(applyDocumentOps(edited.document, created.inverse).isErr()).toBe(true);
    expect(
      findStoryBody(edited.document, { kind: "header", rId: "rIdNew" })?.content.at(0),
    ).toEqual(paragraph("00000006", "Znew"));
  });

  test("lifecycle inverse preserves intervening edits to unowned body fields and stories", () => {
    for (const op of [
      { type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS, sectionIndex: 0, patch: { marginTop: 720 } },
      {
        type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
        sectionIndex: 0,
        story: { kind: "header", rId: "rIdNew" },
        referenceType: "default",
        content: [paragraph("00000006", "new")],
      },
    ] as const satisfies readonly DocumentOp[]) {
      const document = seedDocument("body");
      const changed = expectExact(document, op);
      const edits = [
        {
          type: DOCUMENT_OP_TYPES.INSERT_TEXT,
          at: { story: "main", blockId: "00000001", offset: 0 },
          text: "M",
          runProps: INHERIT_RUN_PROPS,
        },
        {
          type: DOCUMENT_OP_TYPES.INSERT_TEXT,
          at: { story: { kind: "footer", rId: "rIdFooter" }, blockId: "00000003", offset: 0 },
          text: "F",
          runProps: INHERIT_RUN_PROPS,
        },
        ...(op.type === DOCUMENT_OP_TYPES.SET_SECTION_PROPS
          ? [
              {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: { story: { kind: "header", rId: "rIdHeader" }, blockId: "00000002", offset: 0 },
                text: "H",
                runProps: INHERIT_RUN_PROPS,
              } as const,
            ]
          : []),
      ] as const satisfies readonly DocumentOp[];
      const intervening = applyDocumentOps(changed.document, edits).unwrap();
      const undone = applyDocumentOps(
        intervening.document,
        JSON.parse(JSON.stringify(changed.inverse)),
      ).unwrap();
      expect(undone.document).toEqual(applyDocumentOps(document, edits).unwrap().document);
      expect(undone.document.package.document.content).toBe(
        intervening.document.package.document.content,
      );
      expect(undone.document.package.footers).toBe(intervening.document.package.footers);
    }
  });

  test("lifecycle inverse size is independent of unchanged body and secondary-story content", () => {
    const inverses = [1, 128].map((size) => {
      const document = seedDocument("body");
      const content = Array.from({ length: size }, (_, index) =>
        paragraph((index + 16).toString(16).padStart(8, "0"), "body"),
      );
      document.package.document.content = content;
      document.package.document.sections = [{ properties: {}, content }];
      const footer = document.package.footers?.get("rIdFooter");
      if (!footer) return panic("Generated footer must exist.");
      footer.content = [paragraph("00000003", "footer".repeat(size))];
      return (
        [
          { type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS, sectionIndex: 0, patch: { marginTop: 720 } },
          {
            type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
            sectionIndex: 0,
            story: { kind: "header", rId: "rIdNew" },
            referenceType: "default",
            content: [paragraph("00000006", "new")],
          },
        ] as const satisfies readonly DocumentOp[]
      ).map((op) => {
        const edit = applyDocumentOp(document, op).unwrap();
        expect(edit.document.package.document.content).toBe(content);
        expect(edit.touched.modified).toEqual([]);
        for (const inverse of edit.inverse) {
          if (inverse.type !== DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS)
            return panic("Lifecycle edits must produce lifecycle inverses.");
          expect(inverse.parts.body?.content).toBeUndefined();
          expect(inverse.parts.footers).toBeUndefined();
          expect(inverse.parts.footnotes).toBeUndefined();
          expect(inverse.parts.endnotes).toBeUndefined();
        }
        return JSON.stringify(edit.inverse);
      });
    });
    expect(inverses.at(0)).toEqual(inverses.at(1));
  });

  test("removing a reference inside another note preserves its deletion in both note kinds", () => {
    for (const kind of ["footnote", "endnote"] as const) {
      const document = seedDocument("body");
      const holder: Paragraph = {
        type: "paragraph",
        paraId: "00000006",
        content: [
          {
            type: "run",
            formatting: { bold: true },
            content: [
              { type: "text", text: "a" },
              { type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id: 1 },
              { type: "text", text: "b" },
            ],
          },
        ],
      };
      if (kind === "footnote")
        document.package.footnotes?.push({ type: kind, id: 2, content: [holder] });
      else document.package.endnotes?.push({ type: kind, id: 2, content: [holder] });
      const normalized = normalizeForOps(document);
      const removed = expectExact(normalized, {
        type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
        at: { story: { kind, id: 2 }, blockId: "00000006", offset: 1 },
        story: { kind, id: 1 },
      });
      expect(findStoryBody(removed.document, { kind, id: 1 })).toBeUndefined();
      const retained = findStoryBody(removed.document, { kind, id: 2 })?.content.at(0);
      if (retained?.type !== "paragraph") return panic("Retained note must hold its paragraph.");
      const leaves = retained.content.flatMap((inline) =>
        inline.type === "run" ? inline.content : [],
      );
      expect(
        leaves.filter((leaf) => leaf.type === "footnoteRef" || leaf.type === "endnoteRef"),
      ).toEqual([]);
      expect(leaves.map((leaf) => (leaf.type === "text" ? leaf.text : "")).join("")).toBe("ab");
      expect(removed.document.package.document.content).toBe(normalized.package.document.content);
    }
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
