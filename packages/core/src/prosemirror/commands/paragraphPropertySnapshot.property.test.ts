import { expect, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { paragraphFormattingWithAuthoredIndentation } from "../../docx/paragraphPropertySource";
import type { Document } from "../../types/document";
import { expectParagraphAttrs } from "../attrs";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import {
  directParagraphIndentation,
  paragraphIndentationFromFormatting,
  withDirectParagraphIndentation,
} from "../paragraphIndentation";
import { createDocumentStylesPlugin } from "../plugins/documentStyles";
import { acceptChange, rejectChange } from "./comments";
import { paragraphPropertiesSnapshot } from "./propertyChangeScope";

const distance = fc.option(fc.oneof(fc.constant(0), fc.integer({ min: -720, max: 720 })), {
  nil: undefined,
});
const indentation = fc
  .record({
    indentLeft: distance,
    indentRight: distance,
    firstLine: fc.oneof(
      fc.constant(undefined),
      fc.record({
        value: fc.oneof(fc.constant(0), fc.integer({ min: 1, max: 720 })),
        flag: fc.constantFrom("absent", "firstLine", "hanging"),
      }),
    ),
  })
  .map(({ indentLeft, indentRight, firstLine }) =>
    Object.assign(
      {},
      indentLeft === undefined ? {} : { indentLeft },
      indentRight === undefined ? {} : { indentRight },
      Object.assign(
        {},
        firstLine === undefined
          ? {}
          : { indentFirstLine: firstLine.flag === "hanging" ? -firstLine.value : firstLine.value },
        firstLine === undefined || firstLine.flag === "absent"
          ? {}
          : { hangingIndent: firstLine.flag === "hanging" },
      ),
    ),
  );

const authoredIndentation = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Snapshot fixture lost its paragraph.");
  return paragraphIndentationFromFormatting(paragraphFormattingWithAuthoredIndentation(paragraph));
};

// Earlier list matrices exercised effective indentation but their snapshots
// never compared omitted first-line flags with materialized provenance.
test(
  "paragraph snapshots and resolution preserve authored indentation through provenance materialization",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        indentation,
        indentation,
        fc.boolean(),
        async (direct, inherited, captured) => {
          const fixture = {
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "12345678",
                    formatting: { styleId: "Baseline", ...direct },
                    content: [{ type: "run", content: [{ type: "text", text: "Body" }] }],
                  },
                ],
              },
              styles: { styles: [{ styleId: "Baseline", type: "paragraph", pPr: inherited }] },
            },
          } satisfies Document;
          const document = captured
            ? await parseDocx(await createDocx(fixture), {
                preloadFonts: false,
                detectVariables: false,
              })
            : fixture;
          const prose = toProseDoc(document, { styles: document.package.styles });
          const node = prose.firstChild;
          if (node === null) panic("Snapshot fixture has no paragraph.");
          const attrs = expectParagraphAttrs(node);
          const before = paragraphPropertiesSnapshot(node);
          const materialized = node.type.create(
            {
              ...attrs,
              _originalFormatting:
                withDirectParagraphIndentation(
                  attrs._originalFormatting,
                  directParagraphIndentation(attrs),
                ) ?? null,
            },
            node.content,
          );
          expect(paragraphPropertiesSnapshot(materialized)).toEqual(before);

          const changed = materialized.type.create(
            {
              ...materialized.attrs,
              indentLeft: 960,
              _originalFormatting: { ...attrs._originalFormatting, indentLeft: 960 },
              _propertyChanges: [
                {
                  type: "paragraphPropertyChange",
                  info: { id: 42, author: "Reviewer" },
                  previousFormatting: before,
                },
              ],
            },
            node.content,
          );
          for (const decision of ["reject", "accept"] as const) {
            let state = EditorState.create({
              doc: prose.type.create(prose.attrs, [changed]),
              plugins: [createDocumentStylesPlugin(document.package.styles)],
            });
            const expected =
              decision === "reject"
                ? authoredIndentation(document)
                : { ...authoredIndentation(document), indentLeft: 960 };
            const command = decision === "reject" ? rejectChange : acceptChange;
            expect(
              command(0, state.doc.content.size)(state, (transaction) => {
                state = state.apply(transaction);
              }),
            ).toBe(true);
            expect(state.doc.firstChild?.attrs["_propertyChanges"]).toBeNull();
            const resolved = fromProseDoc(state.doc, document, {
              stylesheetSource: { type: "package" },
            });
            expect(authoredIndentation(resolved)).toEqual(expected);
            const reopened = await parseDocx(await createDocx(resolved), {
              preloadFonts: false,
              detectVariables: false,
            });
            expect(authoredIndentation(reopened)).toEqual(expected);
          }
        },
      ),
      {
        numRuns: 60,
        examples: [
          [{ indentLeft: 0, indentFirstLine: 0 }, {}, false],
          [{ indentLeft: 480, indentFirstLine: 120 }, {}, false],
        ],
        id: "paragraph snapshots and resolution preserve authored indentation through provenance materialization",
      },
    );
  },
  propertyTestTimeout(30_000),
);
