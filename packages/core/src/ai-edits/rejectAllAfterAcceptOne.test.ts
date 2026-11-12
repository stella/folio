/**
 * Rejecting every change leaves the document as it was before any of them,
 * whatever one change was accepted on the way: suggestions that insert
 * paragraphs before a block and delete it, one of them accepted into an
 * ordinary tracked change, and a later suggested paragraph after any block.
 * Every original paragraph comes back apart, with its own style and direct
 * alignment, in the live reviewer and in a save rejected after reopening.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const MARKDOWN_LINES = [
  "Service Agreement",
  "This agreement is made between the parties named below.",
  "The Supplier delivers the goods on time and in good order.",
  "The Buyer pays each invoice within thirty days.",
  "Signed in two copies.",
] as const;

const MARKDOWN = MARKDOWN_LINES.map((line, index) =>
  index === 0 ? `# ${line}` : line,
).join("\n\n");

const open = async (bytes: Uint8Array | ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });

const fixture = async (): Promise<Uint8Array> => {
  const bytes = (await ensureParaIds(await createDocx(fromMarkdown(MARKDOWN))))
    .docx;
  const reviewer = await open(bytes);
  const commented = apply(reviewer, "direct", [
    {
      id: "c1",
      type: "commentOnBlock",
      blockId: idOf(reviewer, MARKDOWN_LINES[2]),
      comment: { text: "Define good order." },
    },
    {
      id: "c2",
      type: "commentOnBlock",
      blockId: idOf(reviewer, MARKDOWN_LINES[3]),
      comment: { text: "Thirty days is long." },
    },
  ]);
  expect(commented.applied).toHaveLength(2);
  return new Uint8Array(await reviewer.toBuffer());
};

const idOf = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((entry) => entry.text === text);
  if (!block) throw new Error(`No block reads "${text}"`);
  return block.id;
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "suggested",
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      atomic: false,
      operations,
    },
    { undefinedReferences: "refuse" },
  );

/** What a reader sees of each paragraph. */
const project = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map((block) => ({
    text: block.text,
    kind: block.kind,
    styleId: block.styleId,
    headingLevel: block.headingLevel,
    directAlignment:
      "directAlignment" in block ? block.directAlignment : undefined,
  }));

const reopened = async (reviewer: FolioDocxReviewer) =>
  open(await reviewer.toBuffer());

/**
 * The flow up to the last suggestion: `deleted` is the paragraph the second
 * batch deletes and inserts two paragraphs before, `accept` the index of the
 * change accepted after it, `after` the index of the block the last
 * suggestion inserts a right-aligned paragraph after. Null when that change or
 * block does not exist.
 */
const flow = async (
  bytes: Uint8Array,
  deleted: number,
  accept: number,
  after: number,
): Promise<FolioDocxReviewer | null> => {
  const reviewer = await open(bytes);
  const first = apply(reviewer, "suggested", [
    {
      id: "before",
      type: "insertBeforeBlock",
      blockId: idOf(reviewer, MARKDOWN_LINES[1]),
      text: "Period written supplier invoice.",
    },
  ]);
  expect(first.applied).toHaveLength(1);
  const target = idOf(reviewer, MARKDOWN_LINES[deleted] ?? "");
  const second = apply(reviewer, "suggested", [
    { id: "delete", type: "deleteBlock", blockId: target },
    {
      id: "first",
      type: "insertBeforeBlock",
      blockId: target,
      text: "First added clause.",
    },
    {
      id: "second",
      type: "insertBeforeBlock",
      blockId: target,
      text: "Second added clause.",
    },
  ]);
  expect(second.applied).toHaveLength(3);
  const change = reviewer.getChanges().at(accept);
  if (change === undefined) return null;
  reviewer.acceptChange(change);
  const anchor = reviewer.getContent().at(after);
  if (anchor === undefined) return null;
  apply(reviewer, "suggested", [
    {
      id: "after",
      type: "insertAfterBlock",
      blockId: anchor.id,
      text: "Term agreement party invoice.",
      alignment: "right",
    },
  ]);
  return reviewer;
};

describe("reject all after accepting one of several suggestions", () => {
  for (const deleted of [0, 1, 4]) {
    test(`restores every paragraph when the deleted paragraph is ${deleted}`, async () => {
      const bytes = await fixture();
      const original = project(await open(bytes));
      for (let accept = 0; ; accept += 1) {
        const probe = await flow(bytes, deleted, accept, 0);
        if (probe === null) break;
        const blocks = probe.getContent().length;
        for (let after = 0; after < blocks; after += 1) {
          const label = `accept ${accept}, after ${after}`;
          const live = await flow(bytes, deleted, accept, after);
          if (live === null) continue;
          const saved = await reopened(live);
          live.rejectAll();
          expect({ label, blocks: project(await reopened(live)) }).toEqual({
            label,
            blocks: original,
          });
          saved.rejectAll();
          expect({ label, blocks: project(await reopened(saved)) }).toEqual({
            label,
            blocks: original,
          });
        }
      }
    });
  }
});

// TEMPORARY diagnostic dump, removed before review.
const dump = (label: string, reviewer: FolioDocxReviewer) => {
  const state = Reflect.get(reviewer, "state") as {
    doc: {
      forEach: (
        visit: (node: {
          type: { name: string };
          attrs: Record<string, unknown>;
          textContent: string;
          descendants: (
            visit: (node: {
              isText: boolean;
              text?: string;
              marks: {
                type: { name: string };
                attrs: Record<string, unknown>;
              }[];
            }) => void,
          ) => void;
        }) => void,
      ) => void;
    };
  };
  const lines: string[] = [];
  state.doc.forEach((node) => {
    const runs: string[] = [];
    node.descendants((child) => {
      if (!child.isText) return;
      const marks = child.marks
        .filter(
          (mark) =>
            mark.type.name === "insertion" || mark.type.name === "deletion",
        )
        .map(
          (mark) =>
            `${mark.type.name}:${String(mark.attrs["revisionId"])}:${String(mark.attrs["provenance"])}:${String(mark.attrs["suggestionId"])}`,
        );
      runs.push(`[${child.text ?? ""}|${marks.join(",")}]`);
    });
    lines.push(
      `${node.type.name} pPrMark=${JSON.stringify(node.attrs["pPrMark"])} sugg=${JSON.stringify(node.attrs["_suggestedInsert"])} align=${String(node.attrs["alignment"])} style=${String(node.attrs["styleId"])} pc=${JSON.stringify(node.attrs["_propertyChanges"])} ${runs.join("")}`,
    );
  });
  console.log(`DIAG ${label}\n${lines.join("\n")}`);
};

test("DIAG dump the flow", async () => {
  const bytes = await fixture();
  const reviewer = await open(bytes);
  apply(reviewer, "suggested", [
    {
      id: "before",
      type: "insertBeforeBlock",
      blockId: idOf(reviewer, MARKDOWN_LINES[1]),
      text: "Period written supplier invoice.",
    },
  ]);
  const target = idOf(reviewer, MARKDOWN_LINES[0]);
  apply(reviewer, "suggested", [
    { id: "delete", type: "deleteBlock", blockId: target },
    {
      id: "first",
      type: "insertBeforeBlock",
      blockId: target,
      text: "First added clause.",
    },
    {
      id: "second",
      type: "insertBeforeBlock",
      blockId: target,
      text: "Second added clause.",
    },
  ]);
  dump("after suggestions", reviewer);
  console.log(
    "DIAG changes",
    JSON.stringify(reviewer.getChanges().map((c) => [c.type, c.id, c.text])),
  );
  const mark = reviewer
    .getChanges()
    .find((c) => c.type === "paragraphMarkDeleted");
  if (mark) reviewer.acceptChange(mark);
  dump("after accept", reviewer);
  console.log(
    "DIAG content",
    JSON.stringify(reviewer.getContent().map((b) => [b.id, b.text])),
  );
  const anchor = reviewer.getContent().at(2);
  if (anchor)
    apply(reviewer, "suggested", [
      {
        id: "after",
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Term agreement party invoice.",
        alignment: "right",
      },
    ]);
  dump("after insert after", reviewer);
  const saved = await reopened(reviewer);
  dump("saved", saved);
  reviewer.rejectAll();
  dump("live reject all", reviewer);
  saved.rejectAll();
  dump("saved reject all", saved);
});
