import { expect, spyOn, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { validateOpsDocument } from "../contract";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import * as occupancy from "../editorIntentCensus";
import { formatParaId, ID_LEASE_KINDS } from "../idLease";
import * as ids from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (paraId: string, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});
const bookmark = (id: number, name: string): ParagraphContent[] => [
  { type: "bookmarkStart", id, name },
  run("boundary"),
  { type: "bookmarkEnd", id },
];
const point = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 };
const clipboard = (paragraphs: Paragraph[]) => ({
  type: "replaceFragment" as const,
  from: point,
  to: point,
  openStart: 1 as const,
  openEnd: 1 as const,
  paragraphs,
});

test.each([64, 256])(
  "clipboard compilation reads destination markers only on demand in a %i-paragraph package",
  (width) => {
    const document: Document = {
      package: {
        document: {
          content: [
            paragraph("00000001", [run("base")]),
            paragraph("00000002", bookmark(1, "target")),
            paragraph("00000003", bookmark(2, "target_paste1")),
            ...Array.from({ length: width - 3 }, (_, index) =>
              paragraph(formatParaId(index + 4), [run("untouched")]),
            ),
          ],
        },
      },
    };
    validateOpsDocument(document).unwrap();
    const values = spyOn(Object, "values");
    const existing = spyOn(ids, "packageIdentityKeys");
    const reserved = spyOn(ids, "packageReservedIdentityKeys");
    const paragraphs = spyOn(ids, "packageParagraphIds");
    const membership = spyOn(occupancy, "packageIdOccupancy");
    const destinationVisits = () =>
      values.mock.calls.filter(([value]) => value === document.package).length;
    try {
      compileEditorIntent(document, {
        intent: clipboard([paragraph("00001000", [run("paste")])]),
        mode: { type: "editing" },
      }).unwrap();
      expect(destinationVisits()).toBe(0);
      expect(membership.mock.calls).toHaveLength(0);

      const marked = compileEditorIntent(document, {
        intent: clipboard([
          paragraph("00001000", [
            ...bookmark(99, "target"),
            ...bookmark(100, "other"),
            { type: "hyperlink", anchor: "target", children: [run("link")] },
          ]),
        ]),
        mode: { type: "editing" },
      }).unwrap();
      expect(destinationVisits()).toBe(1);
      expect(membership.mock.calls).toHaveLength(0);
      for (const lookup of [existing, reserved, paragraphs])
        expect(
          lookup.mock.calls.filter(([pkg]) => pkg === document.package),
        ).toHaveLength(0);
      const insertion = marked.ops.find(
        (op) => op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT,
      );
      expect(insertion?.slice.content).toContainEqual({
        type: "bookmarkStart",
        id: 3,
        name: "target_paste2",
      });
      expect(insertion?.slice.content).toContainEqual({
        type: "bookmarkEnd",
        id: 4,
      });
      expect(insertion?.slice.content).toContainEqual({
        type: "hyperlink",
        anchor: "target_paste2",
        children: [run("link")],
      });
      const edited = applyDocumentOps(document, marked.ops).unwrap();
      expect(
        applyDocumentOps(edited.document, edited.inverse).unwrap().document,
      ).toEqual(document);
    } finally {
      membership.mockRestore();
      paragraphs.mockRestore();
      reserved.mockRestore();
      existing.mockRestore();
      values.mockRestore();
    }
  },
);

test("clipboard allocation skips occupied IDs and reserved candidates within the lease", () => {
  const revision = (
    id: number,
  ): Extract<ParagraphContent, { type: "insertion" }> => ({
    type: "insertion",
    info: { id, author: "Author" },
    content: [run("reviewed")],
  });
  const control = (id: number): ParagraphContent => ({
    type: "inlineSdt",
    properties: { id, sdtType: "richText" },
    content: [run("control")],
  });
  const document: Document = {
    package: {
      document: {
        content: [
          paragraph("00000001", [run("base")]),
          paragraph("0000000a", [
            {
              ...revision(10),
              resolutionJoins: {
                before: 0,
                after: 0,
                remove: 0,
                deferredRemove: [{ depth: 0, blockers: [12] }],
              },
            },
            control(10),
          ]),
        ],
      },
    },
  };
  const keptFrom = "2026-03-04T05:00:00Z";
  const intent = clipboard([
    paragraph("00000001", [
      {
        type: "bookmarkStart",
        id: 1,
        name: "paste",
        keptCut: { keptFrom, holders: { before: 10, after: 99 } },
      },
      revision(99),
      { type: "bookmarkEnd", id: 1 },
      control(99),
    ]),
    paragraph("00000002", [run("tail")]),
  ]);
  const mode = {
    type: "editing",
    newIds: { revision: [11], control: [11] },
    idLease: {
      blocks: [
        { kind: ID_LEASE_KINDS.PARAGRAPH, start: 10, end: 12 },
        { kind: ID_LEASE_KINDS.REVISION, start: 10, end: 14 },
        { kind: ID_LEASE_KINDS.CONTROL, start: 10, end: 13 },
      ],
    },
  } satisfies EditorIntentMode;
  const membership = spyOn(occupancy, "packageIdOccupancy");
  try {
    const plan = compileEditorIntent(document, { intent, mode }).unwrap();
    expect(
      membership.mock.calls.filter(([pkg]) => pkg === document.package),
    ).toHaveLength(1);
    const edit = applyDocumentOps(document, plan.ops).unwrap();
    expect(ids.packageParagraphIds(edit.document.package)).toContain(
      "0000000B",
    );
    expect(ids.packageIdentityKeys(edit.document.package)).toContain(
      "revision:13",
    );
    expect(ids.packageIdentityKeys(edit.document.package)).toContain(
      "control:12",
    );
    const pasted = edit.document.package.document.content.flatMap((block) =>
      block.type === "paragraph" ? block.content : [],
    );
    expect(pasted).toContainEqual({
      type: "bookmarkStart",
      id: 1,
      name: "paste",
      keptCut: { keptFrom, holders: { after: 13 } },
    });
    expect(
      applyDocumentOps(edit.document, edit.inverse).unwrap().document,
    ).toEqual(document);
    const exhausted = compileEditorIntent(document, {
      intent,
      mode: {
        ...mode,
        idLease: {
          blocks: mode.idLease.blocks.map(({ kind, start }) => ({
            kind,
            start,
            end: 12,
          })),
        },
      },
    });
    expect(exhausted.isErr() && exhausted.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
    );
  } finally {
    membership.mockRestore();
  }
});
