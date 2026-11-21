import { idKey, type RevisionDecision } from "@stll/docx-core/ops";

import { sameOriginalBlocks } from "../ai-edits/tracked-batch-fidelity";
import type { BlockContent, Paragraph } from "../types/document";

export type IsolatedTextChange = {
  before: Paragraph;
  after: Paragraph;
  beforeIds: number[];
  afterIds: number[];
};

type VerifyWholeTextBatchOptions = {
  changes: readonly IsolatedTextChange[];
  resolve: (
    paragraphs: Paragraph[],
    revisionIds: readonly number[],
    decision: RevisionDecision,
  ) => readonly BlockContent[] | undefined;
};

const paragraphFields = new Set(["type", "paraId", "textId", "content"]);
const wrapperFields = new Set(["type", "info", "content", "resolutionJoins"]);
const infoFields = new Set(["id", "author", "date"]);
const runFields = new Set(["type", "formatting", "content"]);
const textFields = new Set(["type", "text"]);
const onlyFields = (record: object, fields: ReadonlySet<string>) =>
  Object.keys(record).every((field) => fields.has(field));

type Inline = Paragraph["content"][number];

const isSingleTextRun = (nodes: readonly Inline[]): boolean => {
  const run = nodes.at(0);
  if (
    nodes.length !== 1 ||
    run?.type !== "run" ||
    !onlyFields(run, runFields) ||
    (run.formatting !== undefined &&
      Object.keys(run.formatting).length !== 0) ||
    run.content.length !== 1
  )
    return false;
  const text = run.content.at(0);
  return (
    text?.type === "text" &&
    typeof text.text === "string" &&
    onlyFields(text, textFields)
  );
};

const isWholeTextWrapper = (
  node: Inline | undefined,
  kind: "insertion" | "deletion",
): node is Extract<Inline, { type: "insertion" | "deletion" }> => {
  if (
    node?.type !== kind ||
    !onlyFields(node, wrapperFields) ||
    !onlyFields(node.info, infoFields) ||
    !isSingleTextRun(node.content)
  )
    return false;
  const seam = node.resolutionJoins;
  return (
    seam === undefined ||
    (seam.before === 0 &&
      seam.after === 0 &&
      seam.remove === 0 &&
      Object.keys(seam).every(
        (key) =>
          key === "before" ||
          key === "after" ||
          key === "remove" ||
          (key === "acceptance" &&
            kind === "insertion" &&
            seam.acceptance === "merge-plain-runs"),
      ))
  );
};

type WholeTextBatchVerdict =
  | { type: "verified" }
  | {
      type: "fallback";
      reason:
        "singleChange" | "shape" | "sharedIdentity" | "resolution" | "content";
    };

/** Zero-depth whole-text changes with distinct IDs cannot join two paragraphs. */
export const verifyWholeTextBatch = ({
  changes,
  resolve,
}: VerifyWholeTextBatchOptions): WholeTextBatchVerdict => {
  if (changes.length < 2) return { type: "fallback", reason: "singleChange" };
  const ids = new Set<number>();
  const paragraphs = new Set<string>();
  const before: Paragraph[] = [];
  const after: Paragraph[] = [];
  for (const change of changes) {
    const deletion = change.after.content.at(0);
    const insertion = change.after.content.at(1);
    if (
      change.before.paraId === undefined ||
      change.before.paraId !== change.after.paraId ||
      !onlyFields(change.before, paragraphFields) ||
      !onlyFields(change.after, paragraphFields) ||
      !isSingleTextRun(change.before.content) ||
      change.after.content.length !== 2 ||
      !isWholeTextWrapper(deletion, "deletion") ||
      !isWholeTextWrapper(insertion, "insertion")
    )
      return { type: "fallback", reason: "shape" };
    const key = idKey(change.before.paraId);
    if (paragraphs.has(key))
      return { type: "fallback", reason: "sharedIdentity" };
    paragraphs.add(key);
    for (const id of [deletion.info.id, insertion.info.id]) {
      if (ids.has(id)) return { type: "fallback", reason: "sharedIdentity" };
      ids.add(id);
    }
    before.push(change.before);
    after.push(change.after);
  }
  const selected = [...ids];
  const retainsParagraphs = (
    blocks: readonly BlockContent[] | undefined,
  ): blocks is readonly Paragraph[] =>
    blocks !== undefined &&
    blocks.length === after.length &&
    blocks.every(
      (block, index) =>
        block.type === "paragraph" &&
        block.paraId === after.at(index)?.paraId &&
        isSingleTextRun(block.content),
    );
  const restored = resolve(after, selected, "reject");
  if (!retainsParagraphs(restored))
    return { type: "fallback", reason: "resolution" };
  if (!sameOriginalBlocks(before, restored, new Set()))
    return { type: "fallback", reason: "content" };
  const accepted = resolve(after, selected, "accept");
  return retainsParagraphs(accepted)
    ? { type: "verified" }
    : { type: "fallback", reason: "resolution" };
};
