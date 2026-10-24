import { panic } from "better-result";
import { expect, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  type EditorIntent,
} from "../editorIntent";
import { paragraphLogicalText } from "../offsets";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";
import { assertFreshIdentityEquivalent } from "./freshIdentityOracle";
import { resolvePendingRunFormatting } from "./runPropertyReviewOracle";

const source = () => {
  const paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [
      {
        type: "run",
        content: [
          { type: "renderedPageBreak" },
          { type: "tab" },
          { type: "renderedPageBreak" },
        ],
      },
    ],
  } satisfies Paragraph;
  const sectionProperties = { pageWidth: 12240, pageHeight: 15840 };
  return {
    package: {
      document: {
        content: [paragraph],
        sections: [{ properties: sectionProperties, content: [paragraph] }],
        finalSectionProperties: sectionProperties,
        comments: [
          {
            id: 7,
            author: "A",
            content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
          },
        ],
      },
      numbering: undefined,
      footnotes: [
        {
          type: "footnote",
          id: 2,
          content: [{ type: "paragraph", paraId: "7FFFFFF1", content: [] }],
        },
      ],
      settings: { defaultTabStop: 720 },
      properties: { title: "synthetic" },
    },
    warnings: ["kept"],
  } satisfies Document;
};

const intentsFor = (document: Document) => {
  const blockId = "00000001";
  const at = { story: OP_STORIES.MAIN, blockId, offset: 0 };
  const paragraph =
    storyParagraphs(document.package.document).at(0)?.paragraph ??
    panic("The source has its first paragraph.");
  if (paragraph.paraId !== blockId)
    panic("The source has its expected paragraph identity.");
  const length = paragraphLogicalText(paragraph).length;
  const splitAllocation = allocateEditorIntentIds(document, {
    type: "splitParagraph",
    at,
  });
  return [
    {
      type: "formatParagraph",
      at,
      patch: { alignment: "center", keepNext: true },
    },
    {
      type: "formatRun",
      from: at,
      to: { ...at, offset: length },
      patch: { bold: true, italic: false },
    },
    { type: "insertAtom", from: at, to: at, atom: { type: "tab" } },
    {
      type: "replaceText",
      from: at,
      to: { ...at, offset: length },
      text: "edited",
    },
    {
      type: "splitParagraph",
      at,
      newBlockId: splitAllocation.newBlockId,
    },
    {
      type: "setList",
      items: [{ at, ilvl: 0 }],
      target: {
        type: "new",
        num: {
          numId: 1000 + splitAllocation.revisionId,
          abstractNumId: 2000 + splitAllocation.revisionId,
        },
        abstractNum: {
          abstractNumId: 2000 + splitAllocation.revisionId,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
        },
      },
    },
  ] as const satisfies readonly EditorIntent[];
};

type ResolveRevisionOptions = {
  document: Document;
  revisionIds: readonly number[];
  decision: RevisionDecision;
};

const resolveRevision = ({
  document,
  revisionIds,
  decision,
}: ResolveRevisionOptions) => {
  if (revisionIds.length === 0) return document;
  return applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...revisionIds],
    decision,
  }).unwrap().document;
};

const restoreThroughJson = (
  document: Document,
  inverse: readonly DocumentOp[],
) => applyDocumentOps(document, JSON.parse(JSON.stringify(inverse))).unwrap();

const compactOperations = (ops: readonly DocumentOp[]) =>
  JSON.stringify(ops, [
    "type",
    "at",
    "from",
    "to",
    "story",
    "blockId",
    "nextBlockId",
    "offset",
    "zeroWidthBefore",
    "depth",
    "revision",
    "id",
    "author",
    "date",
  ]);

type DiagnosticStage =
  | "apply"
  | "resolve-accept"
  | "resolve-reject"
  | "reject-parity"
  | "accept-parity"
  | "direct-inverse"
  | "tracked-inverse"
  | "direct-redo"
  | "tracked-redo"
  | "complete";

for (const intent of intentsFor(source())) {
  test(`editor intent ${intent.type} has direct and accepted seam parity`, () => {
    const original = source();
    const allocation = allocateEditorIntentIds(original, intent);
    const directPlan = compileEditorIntent(original, {
      intent,
      mode: { type: "editing", newIds: allocation.newIds },
    });
    const trackedPlan = compileEditorIntent(original, {
      intent,
      mode: {
        type: "suggesting",
        revision: {
          id: allocation.revisionId,
          author: "Property",
          date: "2026-10-02T00:00:00Z",
        },
        newIds: allocation.newIds,
      },
    });
    const compilerEvidence = {
      intent: intent.type,
      direct: directPlan.isErr()
        ? { status: "refused", reason: directPlan.error.reason }
        : { status: "accepted", ops: compactOperations(directPlan.value.ops) },
      tracked: trackedPlan.isErr()
        ? { status: "refused", reason: trackedPlan.error.reason }
        : { status: "accepted", ops: compactOperations(trackedPlan.value.ops) },
    };
    if (directPlan.isErr() || trackedPlan.isErr()) {
      process.stderr.write(
        `EDITOR_INTENT_DIAGNOSTIC ${JSON.stringify({ stage: "compile-refused", ...compilerEvidence })}\n`,
      );
      return;
    }
    let stage: DiagnosticStage = "apply";
    let revisionIds: readonly number[] = [];
    try {
      const directEdit = applyDocumentOps(
        original,
        directPlan.value.ops,
      ).unwrap();
      const trackedEdit = applyDocumentOps(
        original,
        trackedPlan.value.ops,
      ).unwrap();
      const originalChanges = storyParagraphs(
        original.package.document,
      ).flatMap(({ paragraph }) => paragraph.propertyChanges ?? []);
      const foldedIds = storyParagraphs(
        trackedEdit.document.package.document,
      ).flatMap(({ paragraph }) =>
        (paragraph.propertyChanges ?? []).flatMap((change) => {
          const previous = originalChanges.find(
            ({ info }) => info.id === change.info.id,
          );
          if (previous === undefined || change.info.author !== "Property")
            return [];
          expect(["formatParagraph", "setList"]).toContain(intent.type);
          assertExactModel(
            change.previousFormatting,
            previous.previousFormatting,
          );
          assertExactModel(change.info, {
            ...previous.info,
            author: "Property",
            date: "2026-10-02T00:00:00Z",
          });
          expect(trackedEdit.revisions).not.toContain(change.info.id);
          return [change.info.id];
        }),
      );
      revisionIds = [...trackedEdit.revisions, ...foldedIds];
      const resolveFolded = (document: Document, decision: RevisionDecision) =>
        intent.type === "formatRun"
          ? resolvePendingRunFormatting(document, decision)
          : resolveRevision({ document, revisionIds: foldedIds, decision });
      stage = "resolve-accept";
      const accepted = resolveRevision({
        document: trackedEdit.document,
        revisionIds,
        decision: REVISION_DECISIONS.ACCEPT,
      });
      stage = "resolve-reject";
      const rejected = resolveRevision({
        document: trackedEdit.document,
        revisionIds,
        decision: REVISION_DECISIONS.REJECT,
      });
      stage = "reject-parity";
      if (intent.type === "setList" && intent.target.type === "new") {
        const rejectedPackage = Object.fromEntries(
          Object.entries(rejected.package).filter(
            ([key]) => key !== "numbering",
          ),
        );
        const originalPackage = Object.fromEntries(
          Object.entries(
            resolveFolded(original, REVISION_DECISIONS.REJECT).package,
          ).filter(([key]) => key !== "numbering"),
        );
        assertExactModel(rejectedPackage, originalPackage);
        expect(rejected.package.numbering?.nums).toContainEqual(
          intent.target.num,
        );
        if (intent.target.abstractNum !== undefined)
          expect(rejected.package.numbering?.abstractNums).toContainEqual(
            intent.target.abstractNum,
          );
      } else {
        assertExactModel(
          resolveFolded(rejected, REVISION_DECISIONS.REJECT),
          resolveFolded(original, REVISION_DECISIONS.REJECT),
        );
      }
      stage = "accept-parity";
      assertFreshIdentityEquivalent({
        actual: resolveFolded(accepted, REVISION_DECISIONS.ACCEPT),
        expected: resolveFolded(directEdit.document, REVISION_DECISIONS.ACCEPT),
        original,
        allocated: allocation.newIds,
      });
      stage = "direct-inverse";
      const directUndo = restoreThroughJson(
        directEdit.document,
        directEdit.inverse,
      );
      assertExactModel(directUndo.document, original);
      stage = "tracked-inverse";
      const trackedUndo = restoreThroughJson(
        trackedEdit.document,
        trackedEdit.inverse,
      );
      assertExactModel(trackedUndo.document, original);
      stage = "direct-redo";
      assertExactModel(
        applyDocumentOps(directUndo.document, directUndo.inverse).unwrap()
          .document,
        directEdit.document,
      );
      stage = "tracked-redo";
      assertExactModel(
        applyDocumentOps(trackedUndo.document, trackedUndo.inverse).unwrap()
          .document,
        trackedEdit.document,
      );
      stage = "complete";
    } finally {
      if (stage !== "complete")
        process.stderr.write(
          `EDITOR_INTENT_DIAGNOSTIC ${JSON.stringify({ stage, ...compilerEvidence, revisionIds })}\n`,
        );
    }
  });
}
