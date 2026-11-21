import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { reviewFieldsOf } from "../review";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "../types";

const tail = { type: "paragraph", paraId: "00000003", content: [] } as const;
const control = (content: Paragraph[]) =>
  ({
    type: "blockSdt",
    properties: { sdtType: "group", id: 9 },
    content,
  }) as const;
const containers = {
  body: { terminal: "story", blocks: (content: Paragraph[]) => content },
  cell: {
    terminal: "cell",
    blocks: (content: Paragraph[]) => [
      {
        type: "table",
        rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
      },
      tail,
    ],
  },
  "terminal control": {
    terminal: "story",
    blocks: (content: Paragraph[]) => [control(content)],
  },
  "terminal custom XML": {
    terminal: "story",
    blocks: (content: Paragraph[]) => [
      {
        type: "blockCustomXml",
        openingXml: '<w:customXml w:element="clause">',
        closingXml: "</w:customXml>",
        content,
      },
    ],
  },
  "interior control": {
    terminal: "interior",
    blocks: (content: Paragraph[]) => [control(content), tail],
  },
} as const satisfies Record<
  string,
  {
    terminal: "story" | "cell" | "interior";
    blocks: (content: Paragraph[]) => readonly BlockContent[];
  }
>;

for (const [name, container] of Object.entries(containers))
  test(`${name}: imported mark inverses restore exact source without permitting a new terminal mark`, () => {
    for (const kind of ["ins", "del", "moveTo", "moveFrom"] as const)
      for (const decision of ["accept", "reject"] as const)
        for (const endpoint of ["plain", "section"] as const) {
          const before = {
            type: "paragraph",
            paraId: "00000001",
            content: [],
          } as const;
          const target = {
            type: "paragraph",
            paraId: "00000002",
            formatting: { indentLeft: 720, runProperties: { fontSize: 18 } },
            content: [
              { type: "run", content: [{ type: "text", text: "owned" }] },
            ],
            pPrMark: { kind, info: { id: 7, author: "Source" } },
            ...(endpoint === "section"
              ? { sectionProperties: { pageWidth: 11906 } }
              : {}),
          } satisfies Paragraph;
          const source = {
            package: {
              document: { content: container.blocks([before, target]) },
            },
          } satisfies Document;
          const { pPrMark: _resolved, ...marked } = target;
          // A rejected added mark takes the section it ended with it.
          const { sectionProperties: _added, ...unsectioned } = marked;
          const unmarked =
            decision === "reject" && (kind === "ins" || kind === "moveTo")
              ? unsectioned
              : marked;
          const expected = {
            package: {
              document: { content: container.blocks([before, unmarked]) },
            },
          } satisfies Document;
          const edit = applyDocumentOp(source, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: [7],
            decision,
          }).unwrap();
          expect(edit.document).toEqual(expected);
          const snapshot =
            container.terminal === "story" ||
            (container.terminal === "cell" &&
              (kind === "moveTo" || kind === "moveFrom"));
          const inverse: DocumentOp[] = JSON.parse(
            JSON.stringify(edit.inverse),
          );
          const undone = applyDocumentOps(edit.document, inverse).unwrap();
          expect(undone.document).toEqual(source);
          const redo: DocumentOp[] = JSON.parse(JSON.stringify(undone.inverse));
          expect(
            applyDocumentOps(undone.document, redo).unwrap().document,
          ).toEqual(expected);
          expect(edit.inverse.at(0)?.type).toBe(
            snapshot
              ? DOCUMENT_OP_TYPES.REPLACE_BLOCKS
              : DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          );
          const ordinary = applyDocumentOp(edit.document, {
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
            story: OP_STORIES.MAIN,
            blockId: target.paraId,
            expected: reviewFieldsOf(unmarked),
            review: reviewFieldsOf(target),
          });
          if (snapshot) {
            expect(ordinary.isErr() && ordinary.error.reason).toBe(
              DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
            );
            const changed = applyDocumentOp(edit.document, {
              type: DOCUMENT_OP_TYPES.INSERT_TEXT,
              at: { story: OP_STORIES.MAIN, blockId: target.paraId, offset: 0 },
              text: "changed",
              runProps: INHERIT_RUN_PROPS,
            }).unwrap().document;
            const beforeStale = structuredClone(changed);
            const stale = applyDocumentOps(changed, inverse);
            expect(stale.isErr() && stale.error.reason).toBe(
              DOCUMENT_OP_REFUSAL_REASONS.STALE,
            );
            expect(changed).toEqual(beforeStale);
          } else {
            // Setting the mark back restores it, not a section it took.
            expect(ordinary.isOk() && ordinary.value.document).toEqual({
              package: {
                document: {
                  content: container.blocks([
                    before,
                    { ...unmarked, pPrMark: target.pPrMark },
                  ]),
                },
              },
            });
          }
        }
  });
