import { panic } from "better-result";
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { serializeDocumentToDocx } from "../../serialize/docx";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import { allocateEditorIntentIds } from "../editorIntent";
import { identityKeysIn } from "../ids";
import { isTrackedWrapper } from "../review";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const sourceKinds = ["insertion", "deletion", "moveFrom", "moveTo"] as const;
const boundaries = [
  { type: "bookmarkStart", id: 9, name: "_Ref9" },
  { type: "bookmarkEnd", id: 9 },
  { type: "moveFromRangeStart", id: 9, name: "move9", author: "Earlier" },
  { type: "moveFromRangeEnd", id: 9 },
  { type: "moveToRangeStart", id: 9, name: "move9", author: "Earlier" },
  { type: "moveToRangeEnd", id: 9 },
] as const satisfies readonly ParagraphContent[];

type FixtureOptions = {
  kind: (typeof sourceKinds)[number];
  side: "before" | "after";
  container: "paragraph" | "wrapper";
  boundary: (typeof boundaries)[number];
  payload: "text" | "equation";
  source: "one-record" | "independent-records";
};

const fixture = ({
  kind,
  side,
  container,
  boundary,
  payload,
  source,
}: FixtureOptions) => {
  const positive =
    payload === "equation"
      ? ({
          type: "mathEquation",
          display: "inline",
          ommlXml: "<m:oMath/>",
        } satisfies ParagraphContent)
      : ({
          type: "run",
          content: [{ type: "text", text: "a" }],
        } satisfies ParagraphContent);
  const markerPiece = {
    type: kind,
    info: { id: 1, author: "Earlier" },
    content: [boundary],
  };
  const positivePiece = {
    type: kind,
    info: { id: 3, author: "Earlier" },
    content: [positive],
  };
  const independent =
    side === "before"
      ? [markerPiece, positivePiece]
      : [positivePiece, markerPiece];
  const content =
    source === "one-record"
      ? [
          {
            type: kind,
            info: { id: 1, author: "Earlier" },
            content:
              side === "before" ? [boundary, positive] : [positive, boundary],
          },
        ]
      : independent;
  const surrounded = [
    {
      type: "preservedInline",
      xml: "<w:proofErr w:type='spellStart'/>",
      text: "",
    },
    ...content,
    {
      type: "preservedInline",
      xml: "<w:proofErr w:type='spellEnd'/>",
      text: "",
    },
  ] satisfies ParagraphContent[];
  const document = normalizeForOps({
    package: {
      properties: {
        created: new Date("2026-01-01T00:00:00Z"),
        modified: new Date("2026-01-01T00:00:00Z"),
      },
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content:
              container === "paragraph"
                ? surrounded
                : [
                    {
                      type: "inlineWrapper",
                      kind: "bidi",
                      control: "embedding",
                      content: surrounded,
                    },
                  ],
          },
        ],
      },
    },
  } satisfies Document);
  const at = {
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    offset: side === "before" ? 0 : 1,
    zeroWidthBefore: side === "before" ? 2 : 0,
  };
  const allocation = allocateEditorIntentIds(document, {
    type: "replaceText",
    from: at,
    to: at,
    text: "typed",
  });
  const operation = {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at,
    slice: {
      openStart: 0,
      openEnd: 0,
      content: [{ type: "run", content: [{ type: "text", text: "typed" }] }],
    },
    revision: {
      id: allocation.revisionId,
      author: "Current",
      date: "2026-11-16T12:00:00Z",
    },
    newIds: allocation.newIds,
  } as const satisfies DocumentOp;
  return { document, operation };
};

const bytes = async (document: Document) =>
  new Uint8Array(await serializeDocumentToDocx(document));

const applyWithExactInverse = async (
  document: Document,
  operation: DocumentOp,
) => {
  const applied = applyDocumentOp(document, operation);
  if (applied.isErr()) throw applied.error;
  const undo = applyDocumentOps(
    applied.value.document,
    JSON.parse(JSON.stringify(applied.value.inverse)),
  );
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
  expect(await bytes(undo.value.document)).toEqual(await bytes(document));
  const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
  if (redo.isErr()) throw redo.error;
  expect(redo.value.document).toStrictEqual(applied.value.document);
  return applied.value;
};

const insertedRevision = (document: Document, id: number) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph")
    return panic("The edited fixture paragraph must exist.");
  const insertion = paragraph.content.find(
    (node) => isTrackedWrapper(node) && node.info.id === id,
  );
  if (insertion === undefined || !isTrackedWrapper(insertion))
    return panic("The inserted tracked record must exist.");
  return insertion;
};

const cases = sourceKinds.flatMap((kind) =>
  (["before", "after"] as const).flatMap((side) =>
    (["paragraph", "wrapper"] as const).map((container) => ({
      kind,
      side,
      container,
    })),
  ),
);

test.each(cases)(
  "rejecting a tracked insertion rejoins its $side marker-only $kind piece in a $container",
  async ({ kind, side, container }) => {
    for (const boundary of boundaries)
      for (const payload of ["text", "equation"] as const) {
        const { document, operation } = fixture({
          kind,
          side,
          container,
          boundary,
          payload,
          source: "one-record",
        });
        const edited = await applyWithExactInverse(document, operation);
        expect(
          insertedRevision(edited.document, operation.revision.id)
            .resolutionJoins?.remove,
        ).toBe(container === "wrapper" ? 2 : 1);
        const ids = identityKeysIn(edited.document.package.document.content);
        expect(new Set(ids).size).toBe(ids.length);
        const rejected = await applyWithExactInverse(edited.document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [operation.revision.id],
          decision: REVISION_DECISIONS.REJECT,
        });
        expect(rejected.document).toStrictEqual(document);
        expect(await bytes(rejected.document)).toEqual(await bytes(document));
        const { revision: _revision, ...directOperation } = operation;
        const direct = await applyWithExactInverse(document, directOperation);
        const accepted = await applyWithExactInverse(edited.document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [operation.revision.id],
          decision: REVISION_DECISIONS.ACCEPT,
        });
        expect(accepted.document).toStrictEqual(direct.document);
        expect(await bytes(accepted.document)).toEqual(
          await bytes(direct.document),
        );
      }
  },
);

test.each(cases)(
  "an independent $side marker-only $kind record in a $container never establishes a cut",
  async ({ kind, side, container }) => {
    for (const boundary of boundaries) {
      const { document, operation } = fixture({
        kind,
        side,
        container,
        boundary,
        payload: "equation",
        source: "independent-records",
      });
      const edited = await applyWithExactInverse(document, operation);
      const insertion = insertedRevision(
        edited.document,
        operation.revision.id,
      );
      expect(insertion.resolutionJoins?.remove).toBe(
        container === "wrapper" ? 1 : 0,
      );
      const rejected = await applyWithExactInverse(edited.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [operation.revision.id],
        decision: REVISION_DECISIONS.REJECT,
      });
      expect(rejected.document).toStrictEqual(document);
    }
  },
);
