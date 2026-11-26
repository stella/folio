import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import { revisionIdsIn } from "../ai-edits/read";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import {
  acceptAllChanges,
  rejectAllChanges,
} from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import {
  compareSectionBoundaryProperties,
  stageSectionBoundaryOperations,
} from "./section-boundary-properties";

const paragraph = (text: string, sectionProperties?: Record<string, unknown>) =>
  schema.node(
    "paragraph",
    sectionProperties === undefined
      ? null
      : { _sectionProperties: sectionProperties },
    text ? schema.text(text) : undefined,
  );

const documentWith = (...paragraphs: ReturnType<typeof paragraph>[]) =>
  schema.node("doc", null, paragraphs);

describe("compareSectionBoundaryProperties", () => {
  test("stages properties only on a retained existing endpoint", () => {
    const current = documentWith(
      paragraph("Before"),
      paragraph("After", { marginLeft: 720 }),
    );
    const target = documentWith(
      paragraph("Before"),
      paragraph("After", { marginLeft: 1440 }),
    );

    expect(compareSectionBoundaryProperties({ current, target })).toEqual({
      status: "matched",
      changes: [
        {
          position: current.child(0).nodeSize,
          current: { marginLeft: 720 },
          target: { marginLeft: 1440 },
        },
      ],
    });
  });

  test("refuses a section endpoint added to retained text", () => {
    const current = documentWith(paragraph("Before"), paragraph("After"));
    const target = documentWith(
      paragraph("Before"),
      paragraph("After", { marginLeft: 1440 }),
    );

    expect(compareSectionBoundaryProperties({ current, target })).toEqual({
      status: "unalignable",
      detail: "section endpoint presence differs",
    });
  });

  test("ignores opaque blocks that cannot carry a section endpoint", () => {
    const current = documentWith(
      paragraph("Before"),
      paragraph("After", { marginLeft: 720 }),
    );
    const target = schema.node("doc", null, [
      paragraph("Before"),
      schema.node("preservedBlock", { xml: '<w:altChunk r:id="rId9"/>' }),
      paragraph("After", { marginLeft: 1440 }),
    ]);

    expect(compareSectionBoundaryProperties({ current, target })).toEqual({
      status: "matched",
      changes: [
        {
          position: current.child(0).nodeSize,
          current: { marginLeft: 720 },
          target: { marginLeft: 1440 },
        },
      ],
    });
  });

  test("ignores block bookmark boundaries between retained endpoints", () => {
    const current = schema.node("doc", null, [
      paragraph("Before"),
      schema.node("blockBookmarkBoundary", {
        type: "start",
        id: 1,
        name: "clause",
      }),
      paragraph("After", { marginLeft: 720 }),
      schema.node("blockBookmarkBoundary", { type: "end", id: 1 }),
    ]);
    const target = schema.node("doc", null, [
      paragraph("Before"),
      paragraph("After", { marginLeft: 1440 }),
      schema.node("blockBookmarkBoundary", {
        type: "start",
        id: 1,
        name: "clause",
      }),
      schema.node("blockBookmarkBoundary", { type: "end", id: 1 }),
    ]);

    expect(compareSectionBoundaryProperties({ current, target })).toEqual({
      status: "matched",
      changes: [
        {
          position: current.child(0).nodeSize + current.child(1).nodeSize,
          current: { marginLeft: 720 },
          target: { marginLeft: 1440 },
        },
      ],
    });
  });
});

const resolveReference = (doc: PMNode, mode: RevisionDecision) => {
  const state = EditorState.create({ doc });
  let resolved = state;
  (mode === "accept" ? acceptAllChanges() : rejectAllChanges())(
    state,
    (transaction) => {
      resolved = state.apply(transaction);
    },
  );
  return resolved.doc;
};

const documentModelOf = (doc: PMNode) => {
  const paragraphs: PMNode[] = [];
  doc.forEach((node, _offset, index) => {
    paragraphs.push(
      node.type.create(
        { ...node.attrs, paraId: (index + 1).toString(16).padStart(8, "0") },
        node.content,
      ),
    );
  });
  return fromProseDoc(schema.node("doc", null, paragraphs), undefined, {
    stylesheetSource: { type: "package" },
  });
};

const resolveModel = (document: Document, decision: RevisionDecision) => {
  const revisionIds = [...revisionIdsIn(document.package.document.content)];
  return toProseDoc(
    applyDocumentOps(document, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        decision,
        revisionIds,
      },
    ]).unwrap().document,
  );
};

type StageOptions = Omit<
  Parameters<typeof stageSectionBoundaryOperations>[0],
  "document" | "accepted"
> & { doc: PMNode };

const stage = ({ doc, ...options }: StageOptions) => {
  const document = documentModelOf(doc);
  const result = stageSectionBoundaryOperations({
    ...options,
    document,
    accepted: () => resolveModel(document, "accept"),
  });
  if (result.status !== "matched")
    panic("Could not stage one section endpoint", { result });
  expect(result.operations).toHaveLength(1);
  expect(result.nextRevisionId).toBe(options.revisionStamp.idSeed + 1);
  return applyDocumentOps(document, result.operations).unwrap().document;
};

const paragraphFacts = (document: PMNode) => {
  const facts: {
    text: string;
    sectionProperties: ReturnType<
      typeof expectParagraphAttrs
    >["_sectionProperties"];
  }[] = [];
  document.forEach((node) => {
    facts.push({
      text: node.textContent,
      sectionProperties: expectParagraphAttrs(node)._sectionProperties,
    });
  });
  return facts;
};

const expectReviewed = ({
  document,
  accepted,
  rejected,
}: {
  document: Document;
  accepted: PMNode;
  rejected: PMNode;
}) => {
  const pending = toProseDoc(document);
  for (const [decision, expected] of [
    ["accept", accepted],
    ["reject", rejected],
  ] as const) {
    const actual = paragraphFacts(resolveModel(document, decision));
    expect(actual).toEqual(paragraphFacts(expected));
    expect(actual).toEqual(paragraphFacts(resolveReference(pending, decision)));
  }
};

test("stages nothing, and reads no accepted projection, for a target without endpoints", () => {
  const deletion = schema.mark("deletion", {
    revisionId: 4,
    author: "Compare",
  });
  const doc = documentWith(
    paragraph("Before"),
    schema.node("paragraph", null, [
      schema.text("Old", [deletion]),
      schema.text(" text"),
    ]),
  );
  const result = stageSectionBoundaryOperations({
    document: documentModelOf(doc),
    accepted: () =>
      panic("Unexpected accepted projection for a target without endpoints"),
    target: documentWith(paragraph("Before"), paragraph("Rewritten text")),
    originalRevisionIdSeed: 10,
    maxRanges: 1,
    revisionStamp: { idSeed: 30, date: "2026-09-13T00:00:00.000Z" },
    author: "Compare",
    mapTargetProperties: () => panic("Unexpected section mapping"),
  });
  expect(result).toEqual({
    status: "matched",
    nextRevisionId: 30,
    operations: [],
  });
});

test("stages a boundary on an already inserted blank paragraph", () => {
  const inserted = schema.node("paragraph", {
    pPrMark: { kind: "ins", info: { id: 20, author: "Compare" } },
  });
  const doc = documentWith(paragraph("Before"), inserted, paragraph("After"));
  const target = documentWith(
    paragraph("Before"),
    schema.node("paragraph", {
      _sectionProperties: { sectionStart: "nextPage", marginLeft: 1440 },
    }),
    paragraph("After"),
  );
  const pending = stage({
    doc,
    target,
    originalRevisionIdSeed: 10,
    maxRanges: 1,
    revisionStamp: { idSeed: 30, date: "2026-09-13T00:00:00.000Z" },
    author: "Compare",
    mapTargetProperties: ({ target: targetProperties }) => ({
      kind: "inserted",
      target: targetProperties,
    }),
  });
  expectReviewed({
    document: pending,
    accepted: target,
    rejected: documentWith(paragraph("Before"), paragraph("After")),
  });
});

test("aligns an inserted endpoint after a preceding deleted paragraph", () => {
  const deleted = schema.node(
    "paragraph",
    { pPrMark: { kind: "del", info: { id: 9, author: "Compare" } } },
    schema.text("Removed", [
      schema.mark("deletion", { revisionId: 9, author: "Compare" }),
    ]),
  );
  const inserted = schema.node("paragraph", {
    pPrMark: { kind: "ins", info: { id: 20, author: "Compare" } },
  });
  const doc = documentWith(
    deleted,
    paragraph("Before"),
    inserted,
    paragraph("After"),
  );
  const target = documentWith(
    paragraph("Before"),
    schema.node("paragraph", {
      _sectionProperties: { sectionStart: "nextPage" },
    }),
    paragraph("After"),
  );
  const pending = stage({
    doc,
    target,
    originalRevisionIdSeed: 10,
    maxRanges: 1,
    revisionStamp: { idSeed: 30, date: "2026-09-13T00:00:00.000Z" },
    author: "Compare",
    mapTargetProperties: ({ target: targetProperties }) => ({
      kind: "inserted",
      target: targetProperties,
    }),
  });
  expectReviewed({
    document: pending,
    accepted: target,
    rejected: documentWith(
      paragraph("Removed"),
      paragraph("Before"),
      paragraph("After"),
    ),
  });
});

test("stages retained endpoint properties alongside an inline text revision", () => {
  const deletion = schema.mark("deletion", {
    revisionId: 4,
    author: "Compare",
  });
  const insertion = schema.mark("insertion", {
    revisionId: 5,
    author: "Compare",
  });
  const previous = { sectionStart: "nextPage", marginLeft: 720 };
  const doc = documentWith(
    schema.node("paragraph", { _sectionProperties: previous }, [
      schema.text("Old", [deletion]),
      schema.text("New", [insertion]),
    ]),
  );
  const target = documentWith(
    schema.node(
      "paragraph",
      {
        _sectionProperties: { sectionStart: "nextPage", marginLeft: 1440 },
      },
      schema.text("New"),
    ),
  );
  const pending = stage({
    doc,
    target,
    originalRevisionIdSeed: 10,
    revisionStamp: { idSeed: 20, date: "2026-09-13T00:00:00.000Z" },
    author: "Compare",
    maxRanges: 1,
    mapTargetProperties: ({ kind, current, target: targetProperties }) => {
      if (kind === "inserted") return { kind, target: targetProperties };
      if (current === undefined) return null;
      return { kind, previous: current, target: targetProperties };
    },
  });
  expectReviewed({
    document: pending,
    accepted: target,
    rejected: documentWith(paragraph("Old", previous)),
  });
});
