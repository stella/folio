/**
 * An applied operation is a promise the document still saves, and saves as
 * asked.
 *
 * Every public operation that removes, merges, splits or rewrites a paragraph
 * is applied to every body paragraph of a three-section package (first and
 * even headers, a table, section breaks on paragraph marks), in every mode.
 * Whatever the batch reports as applied must then:
 *
 * - save with a full repack and reopen;
 * - reopen as the requested outcome: at once for a direct edit, after
 *   accepting for a tracked or suggested one, and as the original after
 *   rejecting a tracked one;
 * - keep saving after a later, unrelated edit.
 *
 * The outcome is judged against a model of the paragraphs and the section
 * each ends: a section break lives on a paragraph's mark (ECMA-376 Part 1
 * §17.6.18), so removing the mark removes the break and its section runs on
 * into the next one, and splitting a paragraph leaves the break on the half
 * that keeps the mark.
 *
 * An operation the batch skips must leave the document saving unchanged.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import type { BlockContent, Paragraph, SectionProperties } from "../types/document";
import { buildSectionsDocx, SECTIONS_FIXTURE_TEXT } from "./__fixtures__/sections";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditApplyMode } from "./types";

/** One top-level block: a paragraph and the section it ends, or a table. */
type Projected =
  | { kind: "paragraph"; text: string; ends: string | null }
  | { kind: "table"; text: string };

type Projection = { blocks: Projected[]; final: string };

/** Sections are told apart by the header parts they reference. */
const sectionName = (properties: SectionProperties): string =>
  (properties.headerReferences ?? []).map(({ type, rId }) => `${type}:${rId}`).join(" ");

const paragraphText = (paragraph: Paragraph): string => {
  let text = "";
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if ("type" in value && value.type === "text" && "text" in value) {
      text += String(value.text);
      return;
    }
    if ("content" in value && Array.isArray(value.content)) {
      for (const child of value.content) visit(child);
    }
  };
  for (const child of paragraph.content) visit(child);
  return text;
};

const projectBlocks = (blocks: readonly BlockContent[]): Projected[] =>
  blocks.map((block) => {
    if (block.type === "paragraph") {
      return {
        kind: "paragraph",
        text: paragraphText(block),
        ends: block.sectionProperties ? sectionName(block.sectionProperties) : null,
      };
    }
    if (block.type === "table") {
      const cells = block.rows.flatMap((row) =>
        row.cells.flatMap((cell) =>
          cell.content.flatMap((child) =>
            child.type === "paragraph" ? [paragraphText(child)] : [],
          ),
        ),
      );
      return { kind: "table", text: cells.join("|") };
    }
    return { kind: "table", text: block.type };
  });

const project = (reviewer: FolioDocxReviewer): Projection => {
  const body = reviewer.toDocument().package.document;
  const final = body.finalSectionProperties;
  return {
    blocks: projectBlocks(body.content),
    final: final ? sectionName(final) : "",
  };
};

const saveAndReopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: "Reviewer" });

type Transform = (blocks: Projected[], mode: FolioAIEditApplyMode) => Projected[];

/** One operation of a case and what it asks for. */
type Step = {
  operation: (blockId: (text: string) => string) => FolioDocumentOperation;
  outcome: Transform;
};

type Case = { name: string; steps: Step[] };

const indexOfText = (blocks: readonly Projected[], text: string): number => {
  const index = blocks.findIndex((block) => block.kind === "paragraph" && block.text === text);
  if (index === -1) throw new Error(`no paragraph "${text}"`);
  return index;
};

const paragraphAt = (blocks: readonly Projected[], index: number) => {
  const block = blocks[index];
  if (block?.kind !== "paragraph") throw new Error(`block ${String(index)} is not a paragraph`);
  return block;
};

/**
 * Removing a paragraph removes its mark, and with it any break it holds.
 *
 * The one exception is the story's last paragraph as a tracked change: its
 * mark cannot be marked deleted (there is no paragraph after it to join), so
 * accepting leaves it empty.
 */
const deleted =
  (text: string): Transform =>
  (blocks, mode) => {
    const index = indexOfText(blocks, text);
    if (mode !== "direct" && index === blocks.length - 1) {
      return blocks.map((block, at) => (at === index ? { ...block, text: "" } : block));
    }
    return blocks.filter((_, at) => at !== index);
  };

/** The joined paragraph ends with the second paragraph's mark. */
const merged =
  (text: string): Transform =>
  (blocks) => {
    const index = indexOfText(blocks, text);
    const first = paragraphAt(blocks, index);
    const second = paragraphAt(blocks, index + 1);
    return [
      ...blocks.slice(0, index),
      { kind: "paragraph", text: first.text + second.text, ends: second.ends },
      ...blocks.slice(index + 2),
    ];
  };

/** The second half keeps the mark, and with it any break. */
const split =
  (text: string, offset: number): Transform =>
  (blocks) => {
    const index = indexOfText(blocks, text);
    const block = paragraphAt(blocks, index);
    return [
      ...blocks.slice(0, index),
      { kind: "paragraph", text: block.text.slice(0, offset), ends: null },
      { kind: "paragraph", text: block.text.slice(offset + 1), ends: block.ends },
      ...blocks.slice(index + 1),
    ];
  };

const replaced =
  (text: string, replacement: string): Transform =>
  (blocks) =>
    blocks.map((block) =>
      block.kind === "paragraph" && block.text === text ? { ...block, text: replacement } : block,
    );

const deleteStep = (text: string): Step => ({
  operation: (id) => ({ id: `delete ${text}`, type: "deleteBlock", blockId: id(text) }),
  outcome: deleted(text),
});

const splitStep = (text: string): Step => ({
  operation: (id) => ({
    id: `split ${text}`,
    type: "splitBlock",
    blockId: id(text),
    offset: text.indexOf(" "),
    separator: " ",
  }),
  outcome: split(text, text.indexOf(" ")),
});

const BODY_PARAGRAPHS = Object.values(SECTIONS_FIXTURE_TEXT).filter(
  (text) => text !== SECTIONS_FIXTURE_TEXT.cell,
);
/** Paragraphs whose next top-level sibling is a paragraph too. */
const MERGEABLE = BODY_PARAGRAPHS.filter(
  (text) => text !== SECTIONS_FIXTURE_TEXT.twoOpens && text !== SECTIONS_FIXTURE_TEXT.threeCloses,
);

const CASES: Case[] = [
  ...BODY_PARAGRAPHS.map(
    (text): Case => ({ name: `deleteBlock "${text}"`, steps: [deleteStep(text)] }),
  ),
  ...MERGEABLE.map(
    (text): Case => ({
      name: `mergeBlockWithNext "${text}"`,
      steps: [
        {
          operation: (id) => ({ id: "merge", type: "mergeBlockWithNext", blockId: id(text) }),
          outcome: merged(text),
        },
      ],
    }),
  ),
  ...BODY_PARAGRAPHS.map(
    (text): Case => ({
      name: `replaceBlock "${text}"`,
      steps: [
        {
          operation: (id) => ({
            id: "replace",
            type: "replaceBlock",
            blockId: id(text),
            text: "Replaced.",
          }),
          outcome: replaced(text, "Replaced."),
        },
      ],
    }),
  ),
  ...BODY_PARAGRAPHS.map(
    (text): Case => ({ name: `splitBlock "${text}"`, steps: [splitStep(text)] }),
  ),
  {
    name: "deleteTable",
    steps: [
      {
        operation: (id) => ({
          id: "table",
          type: "deleteTable",
          blockId: id(SECTIONS_FIXTURE_TEXT.cell),
        }),
        outcome: (blocks) => blocks.filter((block) => block.kind !== "table"),
      },
    ],
  },
  {
    name: "deleteBlock of both section-ending paragraphs in one batch",
    steps: [
      deleteStep(SECTIONS_FIXTURE_TEXT.oneCloses),
      deleteStep(SECTIONS_FIXTURE_TEXT.twoCloses),
    ],
  },
  {
    name: "splitBlock of one section-ending paragraph and deleteBlock of the other in one batch",
    steps: [
      splitStep(SECTIONS_FIXTURE_TEXT.oneCloses),
      deleteStep(SECTIONS_FIXTURE_TEXT.twoCloses),
    ],
  },
];

const MODES: readonly FolioAIEditApplyMode[] = ["direct", "tracked-changes", "suggested"];

const source = await buildSectionsDocx();

const open = async (): Promise<FolioDocxReviewer> =>
  await FolioDocxReviewer.fromBuffer(source, { author: "Reviewer" });

const original = project(await open());

const apply = (reviewer: FolioDocxReviewer, testCase: Case, mode: FolioAIEditApplyMode) => {
  const blocks = reviewer.getContent();
  const blockId = (text: string): string => {
    const block = blocks.find((candidate) => candidate.text === text);
    if (!block) throw new Error(`no block "${text}"`);
    return block.id;
  };
  const operations = testCase.steps.map((step) => ({ step, operation: step.operation(blockId) }));
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: operations.map(({ operation }) => operation),
  });
  const appliedIds = new Set(result.applied.map(({ id }) => id));
  const applied = operations.filter(({ operation }) => appliedIds.has(operation.id));
  let expected = original.blocks;
  for (const { step } of applied) {
    expected = step.outcome(expected, mode);
  }
  return { applied: applied.length > 0, expected };
};

describe("an applied operation saves as requested", () => {
  test("the fixture has three sections, two of them ended by paragraphs", () => {
    expect(
      original.blocks.filter((block) => block.kind === "paragraph" && block.ends),
    ).toHaveLength(2);
    expect(original.final).not.toBe("");
  });

  for (const testCase of CASES) {
    for (const mode of MODES) {
      test(`${testCase.name}, ${mode}`, async () => {
        const reviewer = await open();
        const { applied, expected: blocks } = apply(reviewer, testCase, mode);
        const expected = { blocks, final: original.final };

        if (!applied || mode === "direct") {
          expect(project(await saveAndReopen(reviewer))).toEqual(expected);
          return;
        }

        if (mode === "tracked-changes") {
          const pending = await saveAndReopen(reviewer);
          const rejecting = await FolioDocxReviewer.fromBuffer(await pending.toBuffer());
          pending.acceptAll();
          expect(project(await saveAndReopen(pending))).toEqual(expected);
          rejecting.rejectAll();
          expect(project(await saveAndReopen(rejecting))).toEqual(original);
        } else {
          // A suggestion stays out of the package until someone accepts it.
          expect(project(await saveAndReopen(reviewer))).toEqual(original);
        }
        reviewer.acceptAll();
        expect(project(await saveAndReopen(reviewer))).toEqual(expected);
      });

      test(`${testCase.name}, ${mode}, then a later edit`, async () => {
        const reviewer = await open();
        const { applied, expected } = apply(reviewer, testCase, mode);
        if (!applied) return;
        reviewer.acceptAll();
        const [first] = reviewer.getContent();
        if (!first) throw new Error("no first block");
        const later = reviewer.applyDocumentOperations({
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: [{ id: "lead", type: "insertBeforeBlock", blockId: first.id, text: "Lead." }],
        });
        expect(later.applied).toHaveLength(1);
        expect(project(await saveAndReopen(reviewer))).toEqual({
          blocks: [{ kind: "paragraph", text: "Lead.", ends: null }, ...expected],
          final: original.final,
        });
      });
    }
  }
});
