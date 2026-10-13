import type { Document } from "../../model/document";
import { expect } from "bun:test";
import { panic } from "better-result";
import { applyDocumentOp } from "../apply";
import { storyBody, storyParagraphs } from "../blocks";
import { childNodes, leafSpans, type InlineNode } from "../leaves";
import { DOCUMENT_OP_TYPES, OP_STORIES, type RevisionDecision, type TextPosition } from "../types";

/** Compare complete formatting outcomes when one review consumes an earlier one. */
export const resolvePendingRunFormatting = (
  document: Document,
  decision: RevisionDecision,
): Document => {
  const revisionIds = new Set<number>();
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (node.type === "run")
        for (const change of node.propertyChanges ?? []) revisionIds.add(change.info.id);
      const children = childNodes(node);
      if (children !== undefined) visit(children);
    }
  };
  for (const { paragraph } of storyParagraphs(document.package.document)) visit(paragraph.content);
  if (revisionIds.size === 0) return document;
  return applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...revisionIds],
    decision,
  }).unwrap().document;
};

type OutsideRunReviewOptions = {
  before: Document;
  after: Document;
  from: TextPosition;
  to: TextPosition;
};

/** Resolving reviews must not hide lost authorship outside the edited range. */
export const assertOutsideRunReviewCoverage = ({
  before,
  after,
  from,
  to,
}: OutsideRunReviewOptions): void => {
  const paragraphs = storyParagraphs(storyBody(before, from.story));
  const start = paragraphs.findIndex(({ paragraph }) => paragraph.paraId === from.blockId);
  const end = paragraphs.findIndex(({ paragraph }) => paragraph.paraId === to.blockId);
  if (start < 0 || end < 0) panic("Formatting oracle lost an addressed paragraph");
  const capture = (document: Document) =>
    storyParagraphs(storyBody(document, from.story)).flatMap(({ paragraph }, index) => {
      const coverage: unknown[] = [];
      for (const span of leafSpans(paragraph.content)) {
        const run = span.ancestors.find((node) => node.type === "run");
        if (run?.type !== "run") continue;
        const records = run.propertyChanges?.map(
          ({ info: { id: _id, ...info }, boundaryJoins: _joins, ...change }) => ({
            ...change,
            info,
          }),
        );
        for (let offset = span.before.offset; offset < span.after.offset; offset++) {
          const selected =
            (index > start || (index === start && offset >= from.offset)) &&
            (index < end || (index === end && offset < to.offset));
          if (!selected)
            coverage.push({
              blockId: paragraph.paraId,
              offset,
              formatting: run.formatting,
              records,
            });
        }
      }
      return coverage;
    });
  expect(capture(after)).toStrictEqual(capture(before));
};
