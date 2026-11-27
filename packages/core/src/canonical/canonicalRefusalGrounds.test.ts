import { expect, test } from "bun:test";

import type { Document, Paragraph } from "@stll/docx-core/model";
import {
  applyDocumentOp,
  compileEditorIntent,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  DocumentOpRefusal,
  type EditorIntent,
  type EditorIntentMode,
  OP_STORIES,
} from "@stll/docx-core/ops";

import {
  CANONICAL_CAPABILITIES,
  CANONICAL_CONDITION,
  CANONICAL_GAP,
  CANONICAL_RULE,
} from "../types/canonicalCapabilities";
import { compilerRefusalSkip } from "./canonicalCompilerRefusal";
import { groundForDocumentOpRefusal } from "./canonicalRefusalGrounds";

const R = DOCUMENT_OP_REFUSAL_REASONS;

test("a moved-on document, an exhausted id space, a link suggestion and an untrackable change keep their own grounds", () => {
  expect(groundForDocumentOpRefusal(R.STALE)).toEqual({
    condition: CANONICAL_CONDITION.stale,
  });
  expect(groundForDocumentOpRefusal(R.SECTIONS_OUT_OF_STEP)).toEqual({
    condition: CANONICAL_CONDITION.stale,
  });
  expect(groundForDocumentOpRefusal(R.NEEDS_NEW_IDS)).toEqual({
    condition: CANONICAL_CONDITION.idSpace,
  });
  expect(groundForDocumentOpRefusal(R.UNRECORDABLE_LINK_SUGGESTION)).toEqual({
    rule: CANONICAL_RULE.linkSuggestionShape,
  });
  expect(groundForDocumentOpRefusal(R.TABLE_SNAPSHOT_PRESENCE)).toEqual({
    rule: CANONICAL_RULE.tableSnapshotPresence,
  });
  expect(
    compilerRefusalSkip(
      new DocumentOpRefusal({
        opType: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
        reason: R.TABLE_SNAPSHOT_PRESENCE,
        message: "An own undefined formatting field does not persist.",
      }),
    ),
  ).toMatchObject({ refusal: { rule: CANONICAL_RULE.tableSnapshotPresence } });
  // `w:ins`, `w:del` and the moves hold run-level content only
  // (CT_RunTrackChange), and a block-level `w:sdt` takes no revision of its
  // own: the document's rule, not a gap.
  expect(groundForDocumentOpRefusal(R.NO_REVISION_RECORD)).toEqual({
    rule: CANONICAL_RULE.revisionContentShape,
  });
  expect(groundForDocumentOpRefusal(R.UNTRACKABLE)).toEqual({
    gap: CANONICAL_GAP.trackedChangeRecords,
  });
});

test("every other document operation refusal is a request the document does not allow", () => {
  const named = new Set<string>([
    R.STALE,
    R.SECTIONS_OUT_OF_STEP,
    R.NEEDS_NEW_IDS,
    R.UNRECORDABLE_LINK_SUGGESTION,
    R.TABLE_SNAPSHOT_PRESENCE,
    R.NO_REVISION_RECORD,
    R.UNTRACKABLE,
  ]);
  const others = Object.values(R).filter((reason) => !named.has(reason));
  expect(others.length).toBeGreaterThan(0);
  for (const reason of others)
    expect(groundForDocumentOpRefusal(reason)).toEqual({
      condition: CANONICAL_CONDITION.invalidRequest,
    });
});

test("the gap a document operation refusal names is in the ledger", () => {
  for (const reason of Object.values(R)) {
    const { gap } = groundForDocumentOpRefusal(reason);
    if (gap !== undefined)
      expect(Object.hasOwn(CANONICAL_CAPABILITIES, gap)).toBe(true);
  }
});

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const emptyParagraph = (): Paragraph => ({ type: "paragraph", content: [] });
const SUGGESTING: EditorIntentMode = {
  type: "suggesting",
  revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
  newIds: { revision: Array.from({ length: 40 }, (_, index) => 91 + index) },
};
const EDITING: EditorIntentMode = { type: "editing" };
const at = (blockId: string, offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

// A block-level `w:sdt` takes no revision of its own: CT_SdtBlock holds no
// `w:ins` or `w:del` around itself, and one recorded on the paragraphs inside
// it would, rejected, leave the empty control standing. A suggested paste or
// move of one is refused under the document's rule; made directly, it applies.
test.each<[string, Document, EditorIntent]>([
  [
    "a suggested paste of a block content control",
    { package: { document: { content: [paragraph("00000001", "before")] } } },
    {
      type: "pasteBlocks",
      from: at("00000001", 3),
      to: at("00000001", 3),
      blocks: [
        {
          type: "blockSdt",
          properties: { sdtType: "richText", id: 7 },
          content: [paragraph("00000002", "held")],
        },
      ],
      openStart: 0,
      openEnd: 0,
    },
  ],
  [
    "a suggested move of a block content control",
    {
      package: {
        document: {
          content: [
            paragraph("00000001", "one"),
            {
              type: "blockSdt",
              properties: { sdtType: "richText", id: 7 },
              content: [paragraph("00000005", "held")],
            },
            paragraph("00000003", "two"),
            paragraph("00000004", "three"),
          ],
        },
      },
    },
    {
      type: "moveBlocks",
      from: at("00000001", 3),
      to: at("00000003", 0),
      target: at("00000004", 2),
      blocks: [
        emptyParagraph(),
        {
          type: "blockSdt",
          properties: { sdtType: "richText", id: 7 },
          content: [
            {
              type: "paragraph",
              content: [
                { type: "run", content: [{ type: "text", text: "held" }] },
              ],
            },
          ],
        },
        emptyParagraph(),
      ],
      openStart: 1,
      openEnd: 1,
    },
  ],
])(
  "%s is refused under the revision content rule",
  (_shape, document, intent) => {
    const suggested = compileEditorIntent(document, {
      intent,
      mode: SUGGESTING,
    });
    if (suggested.isOk())
      throw new TypeError("A block content control has no revision record.");
    expect(suggested.error.reason).toBe(R.NO_REVISION_RECORD);
    expect(groundForDocumentOpRefusal(suggested.error.reason)).toEqual({
      rule: CANONICAL_RULE.revisionContentShape,
    });
    expect(
      compileEditorIntent(document, { intent, mode: EDITING }).isOk(),
    ).toBe(true);
  },
);

test("a suggested deletion that would record nothing is a request that changes nothing", () => {
  // The body's last paragraph keeps its mark and holds only another author's
  // pending deletion: a further deletion records nothing, while the direct
  // deletion removes that content.
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "deletion",
                info: { id: 5, author: "Other", date: "2026-01-02T03:04:05Z" },
                content: [
                  { type: "run", content: [{ type: "text", text: "Gone" }] },
                ],
              },
            ],
          },
        ],
      },
    },
  };
  const blockIds = ["00000001"];
  const suggested = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds,
    revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
    newIds: { revision: [91, 92] },
  });
  if (suggested.isOk())
    throw new TypeError("A deletion that records nothing is refused.");
  expect(suggested.error.reason).toBe(R.NO_CHANGE);
  expect(groundForDocumentOpRefusal(suggested.error.reason)).toEqual({
    condition: CANONICAL_CONDITION.invalidRequest,
  });
  expect(
    applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
      story: OP_STORIES.MAIN,
      blockIds,
    }).isOk(),
  ).toBe(true);
});
