/**
 * Content read in one part of a package and placed into another names, once
 * saved and reopened, the picture and the link it named where it was read.
 *
 * Each case takes the paragraph a part holds its picture and link in (the
 * body's, a header's, a footer's, a footnote's, an endnote's or a comment's)
 * and places it into another part, through an editor intent or a document
 * operation that says which story its content was read in, directly or as a
 * suggestion. Every part names its own picture and link by the same ids, so
 * a placed id that is not re-homed resolves to the destination's own. After
 * a save, full or selective, the reopened destination shows the source's
 * picture bytes and link target, every other part still shows its own, and
 * undo restores the document exactly.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type OpStory,
} from "@stll/docx-core/ops";

import { assertExactModel } from "../../../../test/exactModel";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import type { RelationshipStory } from "../docx/storyRelationshipRehoming";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import {
  linkTarget,
  paragraphsHolding,
  pictureAndLink,
  sourceText,
  STORY_PARTS,
  type StoryPartName,
  storyRelationshipPackage,
} from "./__fixtures__/storyRelationshipPackage";

const read = (buffer: ArrayBuffer) =>
  parseDocx(buffer.slice(0), { preloadFonts: false, detectVariables: false });

const only = (paragraphs: readonly Paragraph[]): Paragraph =>
  paragraphs.length === 1
    ? (paragraphs[0] ?? panic("A paragraph is missing."))
    : panic(`Expected one paragraph, found ${paragraphs.length}.`);

/** An identity no paragraph of the package holds. */
const PLACED_PARA_ID = "5A000001";

type Placement = {
  source: (typeof STORY_PARTS)[number];
  destination: (typeof STORY_PARTS)[number];
  route: "intent" | "operation";
  suggesting: boolean;
  selectiveSave: boolean;
};

const placementArbitrary: fc.Arbitrary<Placement> = fc
  .record({
    source: fc.constantFrom(...STORY_PARTS),
    destination: fc.constantFrom(...STORY_PARTS),
    route: fc.constantFrom("intent" as const, "operation" as const),
    suggesting: fc.boolean(),
    selectiveSave: fc.boolean(),
  })
  .filter(({ source, destination }) => source.name !== destination.name);

/** The paragraph of a story placed content goes after. */
const anchorOf = (document: Document, story: OpStory): string => {
  const anchor = paragraphsHolding(document, story, "").at(0);
  return anchor?.paraId ?? panic("The story holds no identified paragraph.");
};

const stateOf = (session: CanonicalSession, story: OpStory) =>
  EditorState.create({
    schema,
    doc:
      story === OP_STORIES.MAIN
        ? session.projection.doc
        : session.projectStory(story).unwrap().doc,
  });

/** Place `paragraph`, read in `from`, into `to`: the commit and its state. */
const placeInto = (
  session: CanonicalSession,
  document: Document,
  {
    paragraph,
    from,
    to,
    route,
  }: {
    paragraph: Paragraph;
    from: RelationshipStory;
    to: RelationshipStory;
    route: Placement["route"];
  },
): { commit: CanonicalCommit; state: EditorState; story: OpStory } => {
  if (to === "comments") {
    const state = stateOf(session, OP_STORIES.MAIN);
    const projection = session.projection;
    return {
      state,
      story: OP_STORIES.MAIN,
      commit: session
        .prepareOps(
          state,
          [
            {
              type: DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
              id: 0,
              content: [{ ...paragraph, paraId: PLACED_PARA_ID }],
            },
          ],
          projection.selectionAt(state).unwrap(),
          { contentStory: from },
        )
        .unwrap(),
    };
  }
  const state = stateOf(session, to);
  const blockId = anchorOf(document, to);
  if (route === "intent")
    return {
      state,
      story: to,
      commit: session
        .prepareIntents(state, {
          intents: [
            {
              type: "insertParagraphs",
              story: to,
              at: { type: "after", blockId },
              paragraphs: [paragraph],
            },
          ],
          contentStory: from,
        })
        .unwrap(),
    };
  const projection =
    to === OP_STORIES.MAIN
      ? session.projection
      : session.projectStory(to).unwrap();
  return {
    state,
    story: to,
    commit: session
      .prepareOps(
        state,
        [
          {
            type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
            story: to,
            at: { type: "after", blockId },
            blocks: [{ ...paragraph, paraId: PLACED_PARA_ID }],
          },
        ],
        projection.selectionAt(state).unwrap(),
        { contentStory: from },
      )
      .unwrap(),
  };
};

const placements = (tally: Map<string, number>) =>
  fc.asyncProperty(placementArbitrary, async (placement) => {
    const { source, destination, route, suggesting, selectiveSave } = placement;
    const opened = await read(await storyRelationshipPackage());
    const own = new Map(
      STORY_PARTS.map(({ name, story }) => [
        name,
        pictureAndLink(
          only(paragraphsHolding(opened, story, sourceText(name))),
        ),
      ]),
    );
    const session = createCanonicalSession(opened).unwrap();
    // A raw operation states its own review; the intent compiles one.
    if (suggesting && route === "intent")
      session.setMode({ type: "suggesting", author: "Placer" });
    const before = session.captureSaveSnapshot().document;
    const paragraph = structuredClone(
      only(paragraphsHolding(opened, source.story, sourceText(source.name))),
    );
    const { commit, state, story } = placeInto(session, opened, {
      paragraph,
      from: source.story,
      to: destination.story,
      route,
    });
    const after = state.apply(commit.transaction);
    commit.publish().unwrap();
    const saved = await serializeCanonicalSave({
      snapshot: session.captureSaveSnapshot(),
      featureFlags: { selectiveSave },
    });
    const reopened = await read(saved.buffer);
    const expected =
      own.get(source.name) ?? panic("The source part has no picture.");
    const placed = pictureAndLink(
      only(
        paragraphsHolding(reopened, destination.story, sourceText(source.name)),
      ),
    );
    expect(placed.image.src).toBe(expected.image.src);
    expect(placed.href).toBe(linkTarget(source.name));
    for (const { name, story: kept } of STORY_PARTS) {
      // The comment's own paragraph is the one content placed into it replaces.
      if (name === "comment" && destination.name === "comment") continue;
      const held = pictureAndLink(
        only(paragraphsHolding(reopened, kept, sourceText(name))),
      );
      const ownPicture = own.get(name) ?? panic("A part has no picture.");
      expect(held.image.src).toBe(ownPicture.image.src);
      expect(held.href).toBe(linkTarget(name));
    }
    session.prepareUndo(after, story).unwrap().publish().unwrap();
    assertExactModel(session.captureSaveSnapshot().document, before);
    const key = `${source.name}->${destination.name}`;
    tally.set(key, (tally.get(key) ?? 0) + 1);
  });

describe("content placed across parts", () => {
  test(
    "a picture and a link placed from one part into another resolve to their own after a save and reopen, and undo restores the document",
    async () => {
      const tally = new Map<string, number>();
      await assertProperty(placements(tally), {
        id: "a picture and a link placed from one part into another resolve to their own after a save and reopen, and undo restores the document",
        numRuns: 60,
      });
      const names = new Set<StoryPartName>(
        [...tally.keys()].flatMap((key) => {
          const [from, to] = key.split("->");
          return STORY_PARTS.filter(
            ({ name }) => name === from || name === to,
          ).map(({ name }) => name);
        }),
      );
      expect(names.size).toBe(STORY_PARTS.length);
    },
    propertyTestTimeout(240_000),
  );
});
