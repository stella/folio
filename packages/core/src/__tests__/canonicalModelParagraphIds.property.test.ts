import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  documentStories,
  storyBody,
  storyParagraphs,
} from "@stll/docx-core/ops";

import {
  createCanonicalEditorHarness,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  holdsUniqueParagraphIds,
  prepareCanonicalDocumentInput,
  prepareCanonicalDocxInput,
} from "../docx/canonicalSessionInput";
import { getParagraphText } from "../docx/paragraphParser";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { EDITOR_MODES } from "./editorHarness";

setDefaultTimeout(propertyTestTimeout(60_000));

const LEAD = "Lead";
const TYPED = "Q";

/** How a generated paragraph holds its id: none, its own, or an earlier one's. */
const ID_CHOICES = ["none", "own", "repeat"] as const;
type IdChoice = (typeof ID_CHOICES)[number];

const paragraphArbitrary = fc.record({
  text: fc.stringMatching(/^[a-zé東]{1,12}$/u),
  id: fc.constantFrom(...ID_CHOICES),
});
type ParagraphSpec = { text: string; id: IdChoice };

const blockArbitrary = fc.oneof(
  paragraphArbitrary.map((paragraph) => ({
    type: "paragraph" as const,
    paragraph,
  })),
  fc
    .tuple(fc.integer({ min: 1, max: 3 }), fc.integer({ min: 1, max: 3 }))
    .chain(([rows, columns]) =>
      fc.record({
        type: fc.constant("table" as const),
        rows: fc.array(
          fc.array(paragraphArbitrary, {
            minLength: columns,
            maxLength: columns,
          }),
          { minLength: rows, maxLength: rows },
        ),
      }),
    ),
);
type BlockSpec =
  | { type: "paragraph"; paragraph: ParagraphSpec }
  | { type: "table"; rows: ParagraphSpec[][] };

/** A host's model, built in memory with the ids each spec asks for. */
const hostModel = (blocks: readonly BlockSpec[]): Document => {
  const document = createEmptyDocument();
  const held: string[] = [];
  const paragraph = (spec: ParagraphSpec): Paragraph => {
    const ids: Record<IdChoice, () => string | undefined> = {
      none: () => undefined,
      own: () => {
        const own = (0x10000000 + held.length + 1).toString(16).toUpperCase();
        held.push(own);
        return own;
      },
      repeat: () => held.at(-1),
    };
    const id = ids[spec.id]();
    return {
      type: "paragraph",
      ...(id !== undefined && { paraId: id }),
      content: [{ type: "run", content: [{ type: "text", text: spec.text }] }],
    };
  };
  const content: BlockContent[] = [
    paragraph({ text: LEAD, id: "none" }),
    ...blocks.map((block): BlockContent => {
      if (block.type === "paragraph") return paragraph(block.paragraph);
      return {
        type: "table",
        columnWidths: (block.rows[0] ?? []).map(() => 2000),
        rows: block.rows.map((row) => ({
          type: "tableRow",
          cells: row.map((cell) => ({
            type: "tableCell",
            content: [paragraph(cell)],
          })),
        })),
      };
    }),
  ];
  document.package.document.content = content;
  return document;
};

/** Every story paragraph in story order: its id and its text. */
const paragraphsOf = (document: Document) =>
  documentStories(document).flatMap((story) =>
    storyParagraphs(storyBody(document, story)).map(({ paragraph }) => ({
      id: paragraph.paraId,
      text: getParagraphText(paragraph),
    })),
  );

const parsed = async (bytes: ArrayBuffer | Uint8Array) =>
  parseDocx(bytes, { detectVariables: false, preloadFonts: false });

test("a host model opens in a canonical session with the ids its package gets, and keeps them through an edit, a save and a reopen", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.constantFrom(...EDITOR_MODES),
      fc.array(blockArbitrary, { minLength: 0, maxLength: 5 }),
      fc.integer({ min: 0, max: LEAD.length }),
      async (mode, blocks, offset) => {
        const model = hostModel(blocks);
        const source = paragraphsOf(model);
        const opened = (await prepareCanonicalDocumentInput(model)).unwrap();
        // The direct route: the same model written by the host and opened as
        // bytes. Both entry points give each paragraph the same id.
        const direct = await parsed(
          (await prepareCanonicalDocxInput(await createDocx(model))).unwrap(),
        );
        const ids = paragraphsOf(opened);
        expect(ids).toEqual(paragraphsOf(direct));
        // Nothing is lost: every paragraph, its text, and every id a
        // paragraph held first; each paragraph now holds an id of its own.
        expect(ids.map(({ text }) => text)).toEqual(
          source.map(({ text }) => text),
        );
        expect(holdsUniqueParagraphIds(opened)).toBe(true);
        const firstHolders = new Set<string>();
        for (const [index, { id }] of source.entries()) {
          if (id === undefined || firstHolders.has(id)) continue;
          firstHolders.add(id);
          expect(ids[index]?.id).toBe(id);
        }

        const harness = createCanonicalEditorHarness(opened, mode);
        let edited: Document;
        try {
          harness.history.setSelection(1 + offset);
          harness.typeText(TYPED);
          expect(harness.refusals).toEqual([]);
          edited = harness.snapshot();
        } finally {
          harness.dispose();
        }
        const lead = `${LEAD.slice(0, offset)}${TYPED}${LEAD.slice(offset)}`;
        const editedParagraphs = paragraphsOf(edited);
        expect(editedParagraphs.map(({ id }) => id)).toEqual(
          ids.map(({ id }) => id),
        );
        expect(editedParagraphs.at(0)?.text ?? panic("No lead")).toContain(
          TYPED,
        );

        const saved = await saveCanonicalHarnessDocument(edited);
        const reopened = await parsed(saved.bytes);
        const reopenedParagraphs = paragraphsOf(reopened);
        expect(reopenedParagraphs.map(({ id }) => id)).toEqual(
          ids.map(({ id }) => id),
        );
        // Typed text sits where the caret was in direct mode; a tracked
        // insertion reads the same, held in its revision.
        expect(reopenedParagraphs.at(0)?.text).toBe(lead);
        expect(reopenedParagraphs.slice(1).map(({ text }) => text)).toEqual(
          source.slice(1).map(({ text }) => text),
        );
        // Reopening, as a model or as bytes, establishes nothing new.
        expect((await prepareCanonicalDocumentInput(reopened)).unwrap()).toBe(
          reopened,
        );
        const reopenedBytes = (
          await prepareCanonicalDocxInput(saved.bytes)
        ).unwrap();
        expect(
          paragraphsOf(await parsed(reopenedBytes)).map(({ id }) => id),
        ).toEqual(ids.map(({ id }) => id));
      },
    ),
    {
      id: "a host model opens in a canonical session with the ids its package gets, and keeps them through an edit, a save and a reopen",
      numRuns: 25,
    },
  );
});
