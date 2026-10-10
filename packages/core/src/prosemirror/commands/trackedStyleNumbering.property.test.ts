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
import { ListExtension, toggleBulletList } from "../extensions/features/ListExtension";
import { schema } from "../schema";
import { ensureParaIds } from "../../docx/ensureParaIds";
import { FolioDocxReviewer } from "../../ai-edits/headless";
import { acceptAllChanges, rejectAllChanges } from "./comments";

type NumberingSource = "style" | "paragraph" | "style-level";

type SourceDocumentOptions = { numId: number; level: number; numberingSource: NumberingSource };

const numberingStateForSource = ({ numId, level, numberingSource }: SourceDocumentOptions) => {
  const reference = { kind: "reference", numId, ilvl: level } as const;
  switch (numberingSource) {
    case "style":
      return { direct: undefined, inherited: reference };
    case "paragraph":
      return { direct: reference, inherited: undefined };
    case "style-level":
      return { direct: { kind: "levelOnly", ilvl: level } as const, inherited: reference };
    default: {
      const unreachable: never = numberingSource;
      return unreachable;
    }
  }
};

const sourceDocument = (options: SourceDocumentOptions): Document => {
  const { numId, level } = options;
  const { direct } = numberingStateForSource(options);
  return {
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
            formatting: {
              styleId: "Numbered",
              ...(direct === undefined ? {} : { numPr: direct }),
            },
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
  };
};

type ExerciseOptions = {
  numId: number;
  level: number;
  operation: "carry" | "list" | "carry then list";
  decision: "accept" | "reject";
  numberingSource: NumberingSource;
};

const exercise = async ({
  numId,
  level,
  operation,
  decision,
  numberingSource,
}: ExerciseOptions) => {
  const numberingState = numberingStateForSource({ numId, level, numberingSource });
  const document = await parseDocx(
    await createDocx(sourceDocument({ numId, level, numberingSource })),
  );
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
  expect(record?.previousFormatting?.numPr ?? undefined).toEqual(numberingState.direct);
  const resolve = decision === "accept" ? acceptAllChanges : rejectAllChanges;
  expect(
    resolve()(state, (tr) => {
      state = state.apply(tr);
    }),
  ).toBe(true);
  const attrs = expectParagraphAttrs(state.doc.child(1));
  if (decision === "reject") {
    expect(attrs.numPr ?? undefined).toEqual(numberingState.direct);
    expect(attrs.numPrFromStyle ?? undefined).toEqual(numberingState.inherited);
    expect(attrs.listNumFmt).toBe("decimal");
  } else if (operation !== "carry") {
    expect(attrs.numPr?.kind).toBe("reference");
    if (attrs.numPr?.kind !== "reference") throw new Error("The list command lost its reference.");
    expect(attrs.numPr.numId).not.toBe(numId);
    expect(attrs.listIsBullet).toBe(true);
    expect(attrs.listNumFmt).toBe("bullet");
  }
  const saved = fromProseDoc(state.doc, document, { stylesheetSource: { type: "package" } });
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
    const directNumbering = findChild(findChild(target, "w", "pPr"), "w", "numPr");
    if (decision === "reject") {
      const directNumberingKind = numberingState.direct?.kind ?? "undefined";
      switch (directNumberingKind) {
        case "reference":
          expect(directNumbering).not.toBeNull();
          expect(paragraph.formatting?.numPr).toEqual({ kind: "reference", numId, ilvl: level });
          expect(paragraph.formatting?.numPrFromStyle).toBeUndefined();
          break;
        case "levelOnly":
          expect(directNumbering).not.toBeNull();
          if (!directNumbering) throw new Error("Expected direct level-only numbering");
          expect(findChild(directNumbering, "w", "numId")).toBeNull();
          expect(paragraph.formatting?.numPr).toEqual({ kind: "levelOnly", ilvl: level });
          expect(paragraph.formatting?.numPrFromStyle).toEqual(numberingState.inherited);
          break;
        case "undefined":
          expect(directNumbering).toBeNull();
          expect(paragraph.formatting?.numPr).toBeUndefined();
          expect(paragraph.formatting?.numPrFromStyle).toEqual(numberingState.inherited);
          break;
        default: {
          const unreachable: never = directNumberingKind;
          return unreachable;
        }
      }
    } else {
      expect(directNumbering).toBeNull();
    }
    expect(paragraph.formatting?.styleId).toBe(decision === "reject" ? "Numbered" : "Plain");
  } else {
    expect(paragraph.formatting?.numPr?.kind).toBe("reference");
    expect(paragraph.listRendering?.isBullet).toBe(true);
    expect(paragraph.listRendering?.numFmt).toBe("bullet");
  }
  expect(expectParagraphAttrs(toProseDoc(reopened).child(1)).numPr).toEqual(attrs.numPr);
};

type StyleNumberingRequest =
  | "toggle bullet"
  | "toggle number"
  | "change level"
  | "new list"
  | "cancel";

const styleNumberingDocument = (): Document => {
  const document = sourceDocument({ numId: 4, level: 0, numberingSource: "style" });
  const numbering = document.package.numbering;
  if (!numbering) throw new Error("Expected numbering definitions");
  const abstractNum = numbering.abstractNums[0];
  if (!abstractNum) throw new Error("Expected an abstract numbering definition");
  abstractNum.levels = [
    { ilvl: 0, numFmt: "decimal", lvlText: "%1.", start: 1 },
    { ilvl: 1, numFmt: "lowerLetter", lvlText: "%1.%2.", start: 1 },
  ];
  return document;
};

const targetProjection = (reviewer: FolioDocxReviewer) => {
  const block = reviewer.getContent().find(({ text }) => text === "Target");
  if (!block || block.kind === "diagnostic") throw new Error("Expected the target paragraph");
  return block;
};

const assertStyleNumberingRequest = async (request: StyleNumberingRequest) => {
  const { docx } = await ensureParaIds(await createDocx(styleNumberingDocument()));
  let saved: ArrayBuffer;
  if (request === "new list" || request === "cancel") {
    const reviewer = await FolioDocxReviewer.fromBuffer(docx, { author: "Reviewer" });
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "numbering-request",
            type: "setBlockParagraphProperties",
            blockId: targetProjection(reviewer).id,
            properties: {
              numbering:
                request === "new list" ? { kind: "newList", format: "numbered" } : { kind: "none" },
            },
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.status).toBe("committed");
    expect(result.issues).toEqual([]);
    saved = await reviewer.toBuffer();
  } else {
    const document = await parseDocx(docx);
    let state = EditorState.create({
      doc: toProseDoc(document),
      plugins: [
        createDocumentStylesPlugin(document.package.styles),
        createDocumentNumberingPlugin(document.package.numbering),
      ],
    });
    const targetPosition = state.doc.child(0).nodeSize;
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, targetPosition + 1)));
    const runtime = ListExtension().onSchemaReady({ schema });
    const command = (() => {
      switch (request) {
        case "toggle bullet":
          return runtime.commands?.toggleBulletList?.();
        case "toggle number":
          return runtime.commands?.toggleNumberedList?.();
        case "change level":
          return runtime.commands?.increaseListLevel?.();
        default: {
          const unreachable: never = request;
          throw new Error(`Unknown list command: ${unreachable}`);
        }
      }
    })();
    if (!command) throw new Error(`List extension has no command for ${request}`);
    expect(
      command(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);
    saved = await repackDocx(
      fromProseDoc(state.doc, document, { stylesheetSource: { type: "package" } }),
    );
  }

  const reopenedBytes = await parseDocx(saved);
  const reopenedProse = toProseDoc(reopenedBytes);
  const renderedAttrs = expectParagraphAttrs(reopenedProse.child(1));
  const reopenedReviewer = await FolioDocxReviewer.fromBuffer(saved);
  const projection = targetProjection(reopenedReviewer);
  const hasEffectiveReference = projection.listReference !== undefined;
  const hasRenderedMarker = renderedAttrs.listNumFmt !== undefined;
  expect(hasEffectiveReference).toBe(hasRenderedMarker);

  switch (request) {
    case "toggle bullet":
      expect(projection.listReference).toBeDefined();
      expect(projection.displayLabel).toBe("•");
      expect(renderedAttrs.listIsBullet).toBe(true);
      expect(projection.statedNumbering.kind).toBe("reference");
      break;
    case "toggle number":
    case "cancel":
      expect(projection.listReference).toBeUndefined();
      expect(projection.displayLabel).toBeUndefined();
      expect(renderedAttrs.listNumFmt).toBeUndefined();
      expect(projection.statedNumbering.kind).toBe("none");
      break;
    case "change level":
      expect(projection.listReference).toEqual({ numId: 4, level: 1 });
      expect(projection.displayLabel).toBe("1.a.");
      expect(renderedAttrs.listNumFmt).toBe("lowerLetter");
      expect(projection.statedNumbering).toEqual({ kind: "levelOnly", ilvl: 1 });
      break;
    case "new list":
      expect(projection.listReference).toBeDefined();
      expect(projection.listReference?.numId).not.toBe(4);
      expect(projection.listReference?.level).toBe(0);
      expect(projection.displayLabel).toBe("1.");
      expect(renderedAttrs.listNumFmt).toBe("decimal");
      expect(projection.statedNumbering.kind).toBe("reference");
      break;
    default: {
      const unreachable: never = request;
      throw new Error(`Unknown numbering request: ${unreachable}`);
    }
  }
};

test("rejecting a suggested property change keeps style numbering inherited", async () => {
  await exercise({
    numId: 4,
    level: 0,
    operation: "carry",
    decision: "reject",
    numberingSource: "style",
  });
});

test("style-numbered requests keep effective membership aligned with saved rendering", async () => {
  for (const request of [
    "toggle bullet",
    "toggle number",
    "change level",
    "new list",
    "cancel",
  ] as const) {
    await assertStyleNumberingRequest(request);
  }
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
              for (const numberingSource of ["style", "paragraph", "style-level"] as const) {
                await exercise({ numId, level, operation, decision, numberingSource });
              }
            }
          }
        },
      ),
      { numRuns: 20 },
    );
  },
  propertyTestTimeout(30_000),
);
