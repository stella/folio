import { panic } from "better-result";
import { expect, test } from "bun:test";
import { type Command, EditorState } from "prosemirror-state";
import { Step, Transform } from "prosemirror-transform";

import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { resolveWholeStory } from "../internal/wholeStoryRevisionResolution";
import type { BlockContent, Document, Paragraph } from "../types/document";
import {
  acceptAIEditRevision,
  rejectAIEditRevision,
} from "./commands/comments";
import { toProseDoc } from "./conversion/toProseDoc";

const tail = { type: "paragraph", paraId: "10000020", content: [] } as const;
const control = (content: BlockContent[]) =>
  ({
    type: "blockSdt",
    properties: { sdtType: "group", id: 9 },
    content,
  }) as const;
const custom = (content: Paragraph[]) =>
  ({
    type: "blockCustomXml",
    openingXml: '<w:customXml w:element="clause">',
    closingXml: "</w:customXml>",
    content,
  }) as const;
const cell = (content: BlockContent[]) =>
  ({
    type: "table",
    rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
  }) as const;

const containers = {
  body: { owner: "first", blocks: (paragraphs: Paragraph[]) => paragraphs },
  cell: {
    owner: "last",
    blocks: (paragraphs: Paragraph[]) => [cell(paragraphs), tail],
  },
  "control interior": {
    owner: "last",
    blocks: (paragraphs: Paragraph[]) => [control(paragraphs), tail],
  },
  "control final": {
    owner: "first",
    blocks: (paragraphs: Paragraph[]) => [control(paragraphs)],
  },
  "custom interior": {
    owner: "last",
    blocks: (paragraphs: Paragraph[]) => [custom(paragraphs), tail],
  },
  "custom final": {
    owner: "first",
    blocks: (paragraphs: Paragraph[]) => [custom(paragraphs)],
  },
  "control final in cell": {
    owner: "last",
    blocks: (paragraphs: Paragraph[]) => [cell([control(paragraphs)]), tail],
  },
  "custom final in cell": {
    owner: "last",
    blocks: (paragraphs: Paragraph[]) => [cell([custom(paragraphs)]), tail],
  },
  "nested transparent final": {
    owner: "first",
    blocks: (paragraphs: Paragraph[]) => [control([custom(paragraphs)])],
  },
  "nested transparent interior": {
    owner: "last",
    blocks: (paragraphs: Paragraph[]) => [control([custom(paragraphs)]), tail],
  },
} as const;

const applyReviewCommand = (state: EditorState, command: Command) => {
  let resolved = state;
  expect(
    command(state, (tr) => {
      resolved = state.apply(tr);
    }),
  ).toBe(true);
  return resolved;
};

for (const [name, container] of Object.entries(containers))
  test(`${name}: empty added mark survivors follow the kernel in whole and partial review`, () => {
    let resolved = 0;
    let refused = 0;
    for (const length of [2, 3, 4])
      for (const kind of ["ins", "moveTo"] as const)
        for (const endpoint of ["plain", "section"] as const)
          for (const decision of ["accept", "reject"] as const) {
            const paragraphs = Array.from(
              { length },
              (_, index) =>
                ({
                  type: "paragraph",
                  paraId: (index + 1).toString(16).padStart(8, "0"),
                  content: [],
                  pPrMark: {
                    kind,
                    info: { id: index + 1, author: "Reviewer" },
                  },
                  ...(endpoint === "section" && index === length - 1
                    ? { sectionProperties: { pageWidth: 11906 } }
                    : {}),
                }) satisfies Paragraph,
            );
            const document = {
              package: { document: { content: container.blocks(paragraphs) } },
            } satisfies Document;
            const source = toProseDoc(document);
            const revisions = paragraphs.map(({ pPrMark }) => pPrMark.info.id);
            const op = {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: revisions,
              decision,
            } as const;
            const outcome = applyDocumentOp(document, op);
            if (
              decision === "reject" &&
              endpoint === "plain" &&
              container.owner === "first"
            ) {
              // Existing structural inverses refuse retiring an imported
              // marked body terminator. This is not a resolved case.
              expect(outcome.isErr()).toBe(true);
              if (outcome.isOk())
                panic("The terminal structural guard must remain enforced.");
              expect(outcome.error.reason).toBe(
                DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
              );
              const { pPrMark: _resolved, ...survivor } =
                paragraphs.at(0) ??
                panic("At least two source paragraphs exist.");
              const whole = resolveWholeStory({
                doc: source,
                mode: decision,
                styleResolver: null,
                numbering: null,
              });
              expect(whole.resolved.toJSON()).toEqual(
                toProseDoc({
                  package: {
                    document: { content: container.blocks([survivor]) },
                  },
                }).toJSON(),
              );
              refused += 1;
              continue;
            }
            const edit = outcome.unwrap();
            resolved += 1;
            const expected = toProseDoc(edit.document);
            const actualIds: string[] = [];
            expected.descendants((node) => {
              if (
                node.type.name === "paragraph" &&
                node.attrs["paraId"] !== tail.paraId
              )
                actualIds.push(node.attrs["paraId"]);
            });
            expect(actualIds).toEqual(
              decision === "accept"
                ? paragraphs.map(({ paraId }) => paraId)
                : [
                    container.owner === "first"
                      ? "00000001"
                      : length.toString(16).padStart(8, "0"),
                  ],
            );
            const whole = resolveWholeStory({
              doc: source,
              mode: decision,
              styleResolver: null,
              numbering: null,
            });
            expect(whole.failed).toBe(false);
            expect(whole.resolved.toJSON()).toEqual(expected.toJSON());
            const replay = new Transform(source);
            const inverse = [];
            for (const step of whole.steps) {
              const copied = Step.fromJSON(
                source.type.schema,
                JSON.parse(JSON.stringify(step.toJSON())),
              );
              inverse.unshift(copied.invert(replay.doc));
              replay.step(copied);
            }
            expect(replay.doc.toJSON()).toEqual(expected.toJSON());
            for (const step of inverse) replay.step(step);
            expect(replay.doc.toJSON()).toEqual(source.toJSON());
            expect(
              applyDocumentOps(edit.document, edit.inverse).unwrap().document,
            ).toEqual(document);
            for (const order of [revisions, revisions.toReversed()]) {
              let model = document;
              let state = EditorState.create({ doc: source });
              for (const id of order) {
                model = applyDocumentOp(model, {
                  ...op,
                  revisionIds: [id],
                }).unwrap().document;
                const command =
                  decision === "accept"
                    ? acceptAIEditRevision(id)
                    : rejectAIEditRevision(id);
                state = applyReviewCommand(state, command);
                expect(state.doc.toJSON()).toEqual(toProseDoc(model).toJSON());
              }
            }
          }
    expect(refused).toBe(container.owner === "first" ? 6 : 0);
    expect(resolved).toBe(container.owner === "first" ? 18 : 24);
  });
