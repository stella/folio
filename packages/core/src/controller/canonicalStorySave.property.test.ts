/**
 * A canonical session's edits in headers, footers and notes reach the file.
 *
 * Each case draws a package with a header, a footer, a footnote and an
 * endnote, each holding top-level paragraphs, a paragraph in a block content
 * control and paragraphs in a table, then applies public operations to those
 * stories, direct or tracked: a replaced or emptied paragraph, a removed
 * paragraph, an added paragraph, a split and a join. Saving the session,
 * through the selective path and through a full repack, and reopening the
 * file reads every story as the saved document holds it.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  OP_STORIES,
  type OpStory,
  paragraphLogicalText,
  paragraphVisibleText,
  storyParagraphs,
} from "@stll/docx-core/ops";

import { packageDocumentArbitrary } from "../../../../test/generators/packageOperationArbitraries";
import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationStory,
} from "../document-operations";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { BlockContent, Document } from "../types/document";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";

/** The stories the generated package holds besides the body, as an operation names them. */
const STORIES = [
  {
    kind: "header",
    public: { type: "header", relationshipId: "rIdHeader" },
    session: { kind: "header", rId: "rIdHeader" },
  },
  {
    kind: "footer",
    public: { type: "footer", relationshipId: "rIdFooter" },
    session: { kind: "footer", rId: "rIdFooter" },
  },
  {
    kind: "footnote",
    public: { type: "footnote", noteId: 2 },
    session: { kind: "footnote", id: 2 },
  },
  {
    kind: "endnote",
    public: { type: "endnote", noteId: 3 },
    session: { kind: "endnote", id: 3 },
  },
] as const satisfies readonly {
  kind: string;
  public: FolioDocumentOperationStory;
  session: OpStory;
}[];

const OPERATION_KINDS = [
  "replaceInBlock",
  "replaceBlock",
  "deleteBlock",
  "insertAfterBlock",
  "splitBlock",
  "mergeBlockWithNext",
] as const;
type OperationKind = (typeof OPERATION_KINDS)[number];

type EditDraw = {
  story: number;
  kind: OperationKind;
  block: number;
  at: number;
  text: string;
  mode: "direct" | "tracked-changes";
};

const editDrawArbitrary: fc.Arbitrary<EditDraw> = fc.record({
  story: fc.nat(),
  kind: fc.constantFrom(...OPERATION_KINDS),
  block: fc.nat(),
  at: fc.nat(),
  text: fc.constantFrom("", "revised", "the Customer"),
  mode: fc.constantFrom("direct", "tracked-changes"),
});

/** The operation a draw names over the blocks a reader of the story sees. */
const operationFor = (
  blocks: readonly { id: string; text: string }[],
  draw: EditDraw,
): FolioDocumentOperation | null => {
  const block = blocks.at(draw.block % Math.max(blocks.length, 1));
  if (block === undefined) return null;
  const { id: blockId, text } = block;
  switch (draw.kind) {
    case "replaceInBlock": {
      const words = text.split(/\s+/u).filter((word) => word !== "");
      const find = words.at(draw.at % Math.max(words.length, 1));
      return find === undefined
        ? null
        : {
            id: "op",
            type: "replaceInBlock",
            blockId,
            find,
            replace: draw.text,
          };
    }
    case "replaceBlock":
      return { id: "op", type: "replaceBlock", blockId, text: draw.text };
    case "deleteBlock":
      return { id: "op", type: "deleteBlock", blockId };
    case "insertAfterBlock":
      return draw.text === ""
        ? null
        : { id: "op", type: "insertAfterBlock", blockId, text: draw.text };
    case "splitBlock":
      return text.length < 2
        ? null
        : {
            id: "op",
            type: "splitBlock",
            blockId,
            offset: 1 + (draw.at % (text.length - 1)),
          };
    case "mergeBlockWithNext":
      return { id: "op", type: "mergeBlockWithNext", blockId };
    default: {
      const unreachable: never = draw.kind;
      throw new Error(`Unknown operation draw ${String(unreachable)}`);
    }
  }
};

/** Each paragraph of a story, as a reader and the file's markup see it. */
const paragraphsOf = (content: readonly BlockContent[]) =>
  storyParagraphs({ content: [...content] }).map(({ paragraph }) => ({
    visible: paragraphVisibleText(paragraph),
    text: paragraphLogicalText(paragraph),
    mark: paragraph.pPrMark?.kind ?? null,
  }));

/** Every story of a document, paragraph by paragraph. */
const storiesOf = (document: Document) => ({
  body: paragraphsOf(document.package.document.content),
  headers: [...(document.package.headers ?? new Map())]
    .toSorted(([left], [right]) => left.localeCompare(right))
    .map(([rId, part]) => [rId, paragraphsOf(part.content)]),
  footers: [...(document.package.footers ?? new Map())]
    .toSorted(([left], [right]) => left.localeCompare(right))
    .map(([rId, part]) => [rId, paragraphsOf(part.content)]),
  footnotes: (document.package.footnotes ?? []).map((note) => [
    note.id,
    paragraphsOf(note.content),
  ]),
  endnotes: (document.package.endnotes ?? []).map((note) => [
    note.id,
    paragraphsOf(note.content),
  ]),
  comments: (document.package.document.comments ?? []).map((comment) => [
    comment.id,
    paragraphsOf(comment.content),
  ]),
});

/** The package as a host stores it: every paragraph identified. */
const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    const prepared = await prepareCanonicalDocxInput(
      await createDocx(structuredClone(document)),
    );
    return prepared.isErr() ? null : prepared.value.slice().buffer;
  } catch {
    return null;
  }
};

const read = (buffer: ArrayBuffer) =>
  parseDocx(buffer.slice(0), { preloadFonts: false, detectVariables: false });

type Tally = Map<string, number>;
const count = (tally: Tally, key: string) => {
  tally.set(key, (tally.get(key) ?? 0) + 1);
};

/** Apply the drawn edits on a canonical session; the stories each one applied in. */
const edited = (
  document: Document,
  draws: readonly EditDraw[],
  tally: Tally,
) => {
  const created = createCanonicalSession(document);
  if (created.isErr()) return null;
  const session = created.value;
  let main = EditorState.create({ schema, doc: session.projection.doc });
  const stateOf = (story: OpStory) => {
    if (story === OP_STORIES.MAIN) return main;
    const projected = session.projectStory(story);
    return projected.isErr()
      ? null
      : EditorState.create({ schema, doc: projected.value.doc });
  };
  const publish = (commit: CanonicalCommit) => {
    const published = publishCanonicalProjection({
      state:
        commit.selection.anchor.story === OP_STORIES.MAIN
          ? main
          : EditorState.create({ schema, doc: commit.transaction.before }),
      commit,
      session,
    });
    if (published.isErr()) return false;
    main = EditorState.create({ schema, doc: session.projection.doc });
    return true;
  };
  const executor = new CanonicalPublicOperations({
    session,
    getState: stateOf,
    publish,
  });
  for (const draw of draws) {
    const story = STORIES[draw.story % STORIES.length];
    if (story === undefined) continue;
    const state = stateOf(story.session);
    if (state === null) continue;
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const operation = operationFor(
      snapshot.blocks.filter((block) => block.kind !== "diagnostic"),
      draw,
    );
    if (operation === null) continue;
    const result = executor.apply({
      snapshot,
      story: story.public,
      batch: { version: 1, mode: draw.mode, operations: [operation] },
    });
    if (result.applied.length > 0)
      count(tally, `${story.kind}:${operation.type}`);
  }
  return session;
};

const storySaves = (tally: Tally) =>
  fc.asyncProperty(
    packageDocumentArbitrary,
    fc.array(editDrawArbitrary, { minLength: 1, maxLength: 3 }),
    fc.boolean(),
    async (generated, draws, selectiveSave) => {
      const buffer = await packaged(generated);
      if (buffer === null) return;
      const session = edited(await read(buffer), draws, tally);
      if (session === null || session.version === 0) return;
      const saved = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
        featureFlags: { selectiveSave },
      });
      expect(storiesOf(await read(saved.buffer))).toEqual(
        storiesOf(saved.document),
      );
      count(tally, "saved");
    },
  );

const GENERATED_RUNS = 40;
const COVERAGE_RUNS = 60;

describe("canonical story edits reach the saved file", () => {
  test(
    "generated header, footer and note edits save and reopen as the session holds them",
    async () => {
      await assertProperty(storySaves(new Map()), {
        id: "generated header, footer and note edits save and reopen as the session holds them",
        numRuns: GENERATED_RUNS,
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "generated story edits reach every story and edit kind",
    async () => {
      const tally: Tally = new Map();
      await checkCoverageSample(storySaves(tally), COVERAGE_RUNS);
      expect(tally.get("saved") ?? 0).toBeGreaterThanOrEqual(COVERAGE_RUNS / 3);
      for (const { kind } of STORIES)
        expect(
          [...tally.keys()].some((key) => key.startsWith(`${kind}:`)),
        ).toBe(true);
      for (const kind of [
        "replaceInBlock",
        "deleteBlock",
        "insertAfterBlock",
        "splitBlock",
        "mergeBlockWithNext",
      ])
        expect([...tally.keys()].some((key) => key.endsWith(`:${kind}`))).toBe(
          true,
        );
    },
    propertyTestTimeout(120_000),
  );
});
