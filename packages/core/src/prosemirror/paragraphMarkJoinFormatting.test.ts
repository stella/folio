import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { Step, Transform } from "prosemirror-transform";

import { joinAtParagraphMark } from "../__tests__/reference-editor/paragraphRevision";
import { resolveWholeStory } from "../__tests__/reference-editor/wholeStoryRevisionResolution";
import { createStyleEngine } from "../style-engine/styleEngine";
import type { Document, ParagraphContent } from "../types/document";
import { expectParagraphAttrs } from "./attrs";
import { toProseDoc } from "./conversion/toProseDoc";
import { joinedParagraphAttrs } from "./paragraphMarkJoin";

const payloads = {
  run: { type: "run", content: [{ type: "text", text: "a" }] },
  field: {
    type: "complexField",
    instruction: "DATE",
    fieldType: "DATE",
    fieldCode: [
      { type: "run", content: [{ type: "instrText", text: " DATE " }] },
    ],
    fieldResult: [{ type: "run", content: [{ type: "text", text: "a" }] }],
  },
  atom: { type: "run", content: [{ type: "tab" }] },
} satisfies Record<string, ParagraphContent>;

for (const [name, payload] of Object.entries(payloads))
  test(`${name} joins refresh body defaults while keeping mark properties and JSON inverses`, () => {
    for (const inTable of [false, true])
      for (const styled of [false, true])
        for (const kind of ["ins", "del", "moveFrom", "moveTo"] as const) {
          const paragraphs = [
            {
              type: "paragraph",
              paraId: "10000001",
              content: [payload],
              pPrMark: { kind, info: { id: 7, author: "Reviewer" } },
            },
            {
              type: "paragraph",
              paraId: "10000002",
              content: [],
              formatting: {
                ...(styled ? { styleId: "Target" } : {}),
                runProperties: { fontSize: 16 },
              },
            },
          ] satisfies Document["package"]["document"]["content"];
          const document = {
            package: {
              styles: {
                docDefaults: { rPr: { fontSize: 22 } },
                styles: [
                  { styleId: "Normal", type: "paragraph", default: true },
                  {
                    styleId: "Target",
                    type: "paragraph",
                    basedOn: "Normal",
                    rPr: { fontSize: 28 },
                  },
                  { styleId: "Table", type: "table", rPr: { fontSize: 26 } },
                ],
              },
              document: {
                content: inTable
                  ? [
                      {
                        type: "table",
                        formatting: { styleId: "Table" },
                        rows: [
                          {
                            type: "tableRow",
                            cells: [{ type: "tableCell", content: paragraphs }],
                          },
                        ],
                      },
                      { type: "paragraph", content: [] },
                    ]
                  : paragraphs,
              },
            },
          } satisfies Document;
          const source = toProseDoc(document);
          const styleResolver = createStyleEngine(document.package.styles);
          const firstPosition = inTable ? 3 : 0;
          const first = source.nodeAt(firstPosition);
          const next = first && source.nodeAt(firstPosition + first.nodeSize);
          if (!first || !next) panic("Both source paragraphs are projected.");
          // Table overlays already select body defaults on the projection
          // leg, including for an empty cell paragraph.
          const defaultSize = inTable ? 26 : 22;
          const inheritedSize = styled ? 28 : defaultSize;
          expect(
            expectParagraphAttrs(next).defaultTextFormatting?.fontSize,
          ).toBe(inTable ? inheritedSize : 16);
          const mode =
            kind === "ins" || kind === "moveTo" ? "reject" : "accept";
          for (const route of ["individual", "whole-story"] as const) {
            const tr = EditorState.create({ doc: source }).tr;
            if (route === "individual")
              joinAtParagraphMark({
                tr,
                paragraphPos: firstPosition,
                paragraph: first,
                next,
                styleResolver,
              });
            else {
              const result = resolveWholeStory({
                doc: source,
                mode,
                styleResolver,
                numbering: null,
              });
              expect(result.failed).toBe(false);
              for (const step of result.steps) tr.step(step);
              expect(tr.doc.eq(result.resolved)).toBe(true);
            }
            const joined = tr.doc.nodeAt(firstPosition);
            if (!joined) panic("The surviving paragraph retains its position.");
            const attrs = expectParagraphAttrs(joined);
            expect(attrs.defaultTextFormatting?.fontSize).toBe(inheritedSize);
            expect(attrs._originalFormatting).toBe(
              expectParagraphAttrs(next)._originalFormatting,
            );
            expect(attrs._originalFormatting?.runProperties?.fontSize).toBe(16);
            expect(attrs._tableRunFormatting).toBe(
              expectParagraphAttrs(next)._tableRunFormatting,
            );
            const replay = new Transform(source);
            const inverses = [];
            for (const step of tr.steps) {
              const copied = Step.fromJSON(
                source.type.schema,
                JSON.parse(JSON.stringify(step.toJSON())),
              );
              inverses.unshift(copied.invert(replay.doc).toJSON());
              replay.step(copied);
            }
            expect(replay.doc.eq(tr.doc)).toBe(true);
            let undone = replay.doc;
            for (const inverse of inverses) {
              const applied = Step.fromJSON(
                source.type.schema,
                JSON.parse(JSON.stringify(inverse)),
              ).apply(undone);
              if (applied.doc === null)
                panic("The inverse restores its source state.");
              undone = applied.doc;
            }
            expect(undone.eq(source)).toBe(true);
          }
        }
  });

test("joining two emitted-empty paragraphs retains the surviving caret defaults", () => {
  const source = toProseDoc({
    package: {
      document: {
        content: [
          { type: "paragraph", content: [] },
          {
            type: "paragraph",
            content: [],
            formatting: { runProperties: { fontSize: 16 } },
          },
        ],
      },
    },
  });
  const first = source.child(0);
  const next = source.child(1);
  const attrs = joinedParagraphAttrs({
    paragraph: first,
    next,
    styleResolver: createStyleEngine(undefined),
  });
  expect(attrs["defaultTextFormatting"]).toBe(
    next.attrs["defaultTextFormatting"],
  );
  expect(attrs["_originalFormatting"]).toBe(next.attrs["_originalFormatting"]);
  expect(expectParagraphAttrs(next).defaultTextFormatting?.fontSize).toBe(16);
});
