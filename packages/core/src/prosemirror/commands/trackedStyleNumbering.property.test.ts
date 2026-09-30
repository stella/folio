import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { findChild, findChildren, parseXml } from "../../docx/xmlParser";
import { EditorState, TextSelection } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { parseDocx } from "../../docx/parser";
import { createDocx, repackDocx } from "../../docx/rezip";
import type { Document } from "../../types/document";
import { expectParagraphAttrs } from "../attrs";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import { carryParagraphProperties } from "../paragraphPropertyCarry";
import { createDocumentNumberingPlugin, getDocumentNumbering } from "../plugins/documentNumbering";
import { createDocumentStylesPlugin, getDocumentStyleResolver } from "../plugins/documentStyles";
import { createSuggestionModePlugin } from "../plugins/suggestionMode";
import { toggleBulletList } from "../extensions/features/ListExtension";
import { acceptAllChanges, rejectAllChanges } from "./comments";

const sourceDocument = (numId: number, level: number): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          formatting: { styleId: "Plain" },
          content: [{ type: "run", content: [{ type: "text", text: "Source" }] }],
        },
        {
          type: "paragraph",
          formatting: { styleId: "Numbered" },
          content: [{ type: "run", content: [{ type: "text", text: "Target" }] }],
        },
      ],
    },
    styles: {
      styles: [
        { type: "paragraph", styleId: "Plain" },
        {
          type: "paragraph",
          styleId: "Numbered",
          pPr: { numPr: { kind: "reference", numId, ilvl: level } },
        },
      ],
    },
    numbering: {
      nums: [{ numId, abstractNumId: 1 }],
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: level, numFmt: "decimal", lvlText: `%${level + 1}.`, start: 1 }],
        },
      ],
    },
  },
});

type ExerciseOptions = {
  numId: number;
  level: number;
  operation: "carry" | "list" | "carry then list";
  decision: "accept" | "reject";
};

const exercise = async ({ numId, level, operation, decision }: ExerciseOptions) => {
  const document = await parseDocx(await createDocx(sourceDocument(numId, level)));
  let state = EditorState.create({
    doc: toProseDoc(document),
    plugins: [
      createDocumentStylesPlugin(document.package.styles),
      createDocumentNumberingPlugin(document.package.numbering),
      createSuggestionModePlugin(true, "Reviewer"),
    ],
  });
  const position = state.doc.child(0).nodeSize;
  if (operation !== "list") {
    const tr = state.tr;
    carryParagraphProperties({
      tr,
      position,
      source: state.doc.child(0),
      styleResolver: getDocumentStyleResolver(state),
      numbering: getDocumentNumbering(state),
      revision: { id: 42, author: "Reviewer", date: "2026-09-01T00:00:00Z" },
    });
    state = state.apply(tr);
  }
  if (operation !== "carry") {
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, position + 1)));
    expect(
      toggleBulletList(state, (tr) => {
        state = state.apply(tr);
      }),
    ).toBe(true);
  }
  const record = expectParagraphAttrs(state.doc.child(1))._propertyChanges?.at(0);
  expect(record).toBeDefined();
  expect(record?.previousFormatting?.numPr ?? undefined).toBeUndefined();
  const resolve = decision === "accept" ? acceptAllChanges : rejectAllChanges;
  expect(
    resolve()(state, (tr) => {
      state = state.apply(tr);
    }),
  ).toBe(true);
  const attrs = expectParagraphAttrs(state.doc.child(1));
  if (decision === "reject") {
    expect(attrs.numPr).toEqual({ kind: "reference", numId, ilvl: level });
    expect(attrs.listNumFmt).toBe("decimal");
  }
  const saved = fromProseDoc(state.doc, document);
  const bytes = await repackDocx(saved);
  const reopened = await parseDocx(bytes);
  const paragraph = reopened.package.document.content.at(1);
  expect(paragraph?.type).toBe("paragraph");
  if (paragraph?.type !== "paragraph") throw new Error("Expected target paragraph");
  expect(paragraph.propertyChanges).toBeUndefined();
  if (decision === "reject" || operation === "carry") {
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")?.async("string");
    expect(xml).toBeDefined();
    const root = parseXml(xml ?? "");
    const body = findChild(findChild(root, "w", "document"), "w", "body");
    const target = findChildren(body, "w", "p").at(1);
    expect(findChild(findChild(target, "w", "pPr"), "w", "numPr")).toBeNull();
    expect(paragraph.formatting?.styleId).toBe(decision === "reject" ? "Numbered" : "Plain");
  } else {
    expect(paragraph.formatting?.numPr?.kind).toBe("reference");
  }
  expect(expectParagraphAttrs(toProseDoc(reopened).child(1)).numPr).toEqual(attrs.numPr);
};

test("rejecting a suggested property change keeps style numbering inherited", async () => {
  await exercise({ numId: 4, level: 0, operation: "carry", decision: "reject" });
});

test(
  "tracked property records and resolution preserve style numbering provenance",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 1, max: 20 }),
        fc.integer({ min: 0, max: 8 }),
        async (numId, level) => {
          for (const operation of ["carry", "list", "carry then list"] as const) {
            for (const decision of ["accept", "reject"] as const) {
              await exercise({ numId, level, operation, decision });
            }
          }
        },
      ),
      { numRuns: 20 },
    );
  },
  propertyTestTimeout(30_000),
);
