import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  applyBatch,
  captureDocumentOp,
  DOCUMENT_OP_SCHEMA_VERSION,
  OP_STORIES,
  type OpStory,
  paragraphLength,
  type SequencedBatch,
  storyBody,
  storyParagraphs,
  type TextPosition,
} from "@stll/docx-core/ops";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  type CanonicalSecondaryStory,
  type CanonicalSelection,
  type CanonicalSession,
  createCanonicalSession,
} from "../canonical/canonicalSession";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";

setDefaultTimeout(propertyTestTimeout(60_000));

/*
 * A peer edits the body, a header and a footnote while this session holds a
 * selection in each: the body's in its editor, the header's and the note's
 * as their editors recorded them. Every peer batch enters through
 * `prepareRemote`. Each selection end is a position in its paragraph, and a
 * plain model of the paragraphs says where the peer's edit leaves it: text
 * inserted before it pushes it on and text inserted at it lands after it,
 * a deletion covering it leaves it where the deletion was, a split puts it
 * in the half that holds its text, and a join keeps its offset in the
 * joined paragraph.
 */

const HEADER = { kind: "header", rId: "rIdHeader" } as const;
const NOTE = { kind: "footnote", id: 12 } as const;
const SECONDARY: readonly CanonicalSecondaryStory[] = [HEADER, NOTE];
const STORIES: readonly OpStory[] = [OP_STORIES.MAIN, HEADER, NOTE];
/** The body paragraph holding the note reference; the peer leaves it alone. */
const REFERENCE_PARAGRAPH = "0000000A";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const fixture = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("12345678", "Body seed"),
        {
          type: "paragraph",
          paraId: REFERENCE_PARAGRAPH,
          content: [
            {
              type: "run",
              content: [
                { type: "text", text: "Ref" },
                { type: "footnoteRef", id: 12 },
              ],
            },
          ],
          sectionProperties: {
            headerReferences: [{ type: "default", rId: HEADER.rId }],
          },
        },
      ],
    },
    headers: new Map([
      [
        HEADER.rId,
        {
          type: "header",
          hdrFtrType: "default",
          content: [paragraph("23456789", "Header seed")],
        },
      ],
    ]),
    footnotes: [
      {
        type: "footnote",
        id: 12,
        content: [
          {
            ...paragraph("34567890", "Footnote seed"),
            // A note opens with its own reference mark, as a saved note does.
            content: [
              {
                type: "run",
                content: [{ type: "noteMarker", kind: "footnote" }],
              },
              {
                type: "run",
                content: [{ type: "text", text: "Footnote seed" }],
              },
            ],
          },
        ],
      },
    ],
  },
});

type Place = { blockId: string; offset: number };
type Ends = { anchor: Place; head: Place };

/** The paragraphs of `story` the peer may edit: id and length, in order. */
const editable = (document: Document, story: OpStory) =>
  storyParagraphs(storyBody(document, story)).flatMap(({ paragraph: item }) =>
    item.paraId === undefined || item.paraId === REFERENCE_PARAGRAPH
      ? []
      : [{ blockId: item.paraId, length: paragraphLength(item) }],
  );

const pick = <T>(items: readonly T[], seed: number): T =>
  items.at(seed % items.length) ?? panic("Nothing to pick from.");

const placeIn = (document: Document, story: OpStory, seed: number): Place => {
  const target = pick(editable(document, story), seed);
  return {
    blockId: target.blockId,
    offset: Math.floor(seed / 7) % (target.length + 1),
  };
};

const positionOf = (story: OpStory, place: Place): TextPosition => ({
  story,
  ...place,
});

const endsOf = (selection: CanonicalSelection | null): Ends | null =>
  selection === null
    ? null
    : {
        anchor: {
          blockId: selection.anchor.blockId,
          offset: selection.anchor.offset,
        },
        head: {
          blockId: selection.head.blockId,
          offset: selection.head.offset,
        },
      };

/** A state of `story`'s projection with a text selection over `ends`. */
const stateAt = (
  session: CanonicalSession,
  story: OpStory,
  ends?: Ends,
): EditorState => {
  const projection = session.projectStory(story).unwrap();
  const state = EditorState.create({ schema, doc: projection.doc });
  if (ends === undefined) return state;
  const anchor = projection.positionAt(positionOf(story, ends.anchor)).unwrap();
  const head = projection.positionAt(positionOf(story, ends.head)).unwrap();
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, anchor, head)),
  );
};

type PeerEdit =
  | { type: "insert"; at: Place; text: string }
  | { type: "delete"; from: Place; length: number }
  | { type: "split"; at: Place; blockId: string }
  | { type: "join"; blockId: string; nextBlockId: string };

/** Where a selection end goes over the peer's edit, by the paragraph model. */
const moved = (
  place: Place,
  edit: PeerEdit,
  before: Document,
  after: Document,
  story: OpStory,
): Place => {
  switch (edit.type) {
    case "insert":
      return place.blockId === edit.at.blockId && place.offset > edit.at.offset
        ? { ...place, offset: place.offset + edit.text.length }
        : place;
    case "delete": {
      if (place.blockId !== edit.from.blockId) return place;
      const end = edit.from.offset + edit.length;
      if (place.offset <= edit.from.offset) return place;
      return place.offset <= end
        ? { ...place, offset: edit.from.offset }
        : { ...place, offset: place.offset - edit.length };
    }
    case "split": {
      if (place.blockId !== edit.at.blockId) return place;
      // The halves stand where the split paragraph stood, in reading order.
      const index = editable(before, story).findIndex(
        ({ blockId }) => blockId === edit.at.blockId,
      );
      const halves = editable(after, story).slice(index, index + 2);
      const [first, second] = halves;
      if (first === undefined || second === undefined)
        return panic("A split left no two halves.");
      return place.offset <= edit.at.offset
        ? { blockId: first.blockId, offset: place.offset }
        : { blockId: second.blockId, offset: place.offset - edit.at.offset };
    }
    case "join": {
      const paragraphs = editable(before, story);
      const first =
        paragraphs.find(({ blockId }) => blockId === edit.blockId) ??
        panic("A join lost its first paragraph.");
      const index = paragraphs.indexOf(first);
      const survivor =
        editable(after, story).at(index) ?? panic("A join left nothing.");
      if (place.blockId === edit.blockId)
        return { blockId: survivor.blockId, offset: place.offset };
      if (place.blockId === edit.nextBlockId)
        return {
          blockId: survivor.blockId,
          offset: first.length + place.offset,
        };
      return place;
    }
    default: {
      const unreachable: never = edit;
      return unreachable;
    }
  }
};

/** Whether a paragraph of `story` holds the note's own reference mark. */
const holdsMark = (document: Document, story: OpStory, blockId: string) =>
  storyParagraphs(storyBody(document, story)).some(
    ({ paragraph: item }) =>
      item.paraId === blockId &&
      item.content.some(
        (child) =>
          child.type === "run" &&
          child.content.some((leaf) => leaf.type === "noteMarker"),
      ),
  );

/**
 * Where a selection end may go over the peer's edit. The note's own mark is
 * a zero-width child at offset 0 of its paragraph, so an end and an insertion
 * both at that offset are ordered by which side of the mark each stands on,
 * which the paragraph model does not track: either side is right.
 */
const candidates = (
  place: Place,
  edit: PeerEdit,
  before: Document,
  after: Document,
  story: OpStory,
): Place[] => {
  const ahead = moved(place, edit, before, after, story);
  return edit.type === "insert" &&
    edit.at.blockId === place.blockId &&
    edit.at.offset === 0 &&
    place.offset === 0 &&
    holdsMark(before, story, place.blockId)
    ? [ahead, { ...place, offset: edit.text.length }]
    : [ahead];
};

type Step =
  | { type: "peer"; story: number; kind: number; seed: number; text: string }
  | { type: "select"; story: number; anchor: number; head: number };

const stepArbitrary: fc.Arbitrary<Step> = fc.oneof(
  {
    weight: 5,
    arbitrary: fc.record({
      type: fc.constant("peer" as const),
      story: fc.nat({ max: STORIES.length - 1 }),
      kind: fc.nat({ max: 3 }),
      seed: fc.nat(),
      text: fc.constantFrom("x", "Remote ", "éé"),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("select" as const),
      story: fc.nat({ max: STORIES.length - 1 }),
      anchor: fc.nat(),
      head: fc.nat(),
    }),
  },
);

const selectionsArbitrary = fc.array(fc.tuple(fc.nat(), fc.nat()), {
  minLength: STORIES.length,
  maxLength: STORIES.length,
});

const paragraphLengthAt = (
  document: Document,
  story: OpStory,
  blockId: string,
): number =>
  editable(document, story).find((item) => item.blockId === blockId)?.length ??
  panic("The paragraph is gone.");

/** The peer's edit for a step, or undefined when the story offers none. */
const peerEdit = (
  document: Document,
  story: OpStory,
  step: Extract<Step, { type: "peer" }>,
  freshId: () => string,
): PeerEdit | undefined => {
  const at = placeIn(document, story, step.seed);
  switch (step.kind) {
    case 0:
      return { type: "insert", at, text: step.text };
    case 1: {
      const length = paragraphLengthAt(document, story, at.blockId);
      const count = Math.min(1 + (step.seed % 5), length - at.offset);
      return count > 0
        ? { type: "delete", from: at, length: count }
        : undefined;
    }
    case 2:
      // A split at the note's own mark decides which half keeps the mark by
      // the side it names, which the model does not track; the mark stays at
      // the start of its paragraph.
      return at.offset === 0 && holdsMark(document, story, at.blockId)
        ? undefined
        : { type: "split", at, blockId: freshId() };
    default: {
      const paragraphs = editable(document, story);
      const index = paragraphs.findIndex(
        ({ blockId }) => blockId === at.blockId,
      );
      const next = paragraphs.at(index + 1);
      return next === undefined
        ? undefined
        : { type: "join", blockId: at.blockId, nextBlockId: next.blockId };
    }
  }
};

/**
 * The peer makes its edit; returns the batch the room would sequence, or
 * undefined when the peer's session refuses the edit (a note's own mark).
 */
const peerBatch = (
  peer: CanonicalSession,
  story: OpStory,
  edit: PeerEdit,
  revision: number,
): SequencedBatch | undefined => {
  const before = peer.document;
  const intent = (() => {
    switch (edit.type) {
      case "insert":
        return {
          type: "replaceText" as const,
          from: positionOf(story, edit.at),
          to: positionOf(story, edit.at),
          text: edit.text,
        };
      case "delete":
        return {
          type: "replaceText" as const,
          from: positionOf(story, edit.from),
          to: positionOf(story, {
            ...edit.from,
            offset: edit.from.offset + edit.length,
          }),
          text: "",
        };
      case "split":
        return {
          type: "splitParagraph" as const,
          at: positionOf(story, edit.at),
          newBlockId: edit.blockId,
        };
      case "join":
        return {
          type: "joinParagraphs" as const,
          story,
          blockId: edit.blockId,
          nextBlockId: edit.nextBlockId,
        };
      default: {
        const unreachable: never = edit;
        return unreachable;
      }
    }
  })();
  const prepared = peer.prepareIntents(stateAt(peer, story), {
    intents: [intent],
  });
  if (prepared.isErr()) return undefined;
  const commit = prepared.value;
  expect(commit.publish().isOk()).toBe(true);
  const applied = applyBatch(before, commit.ops).unwrap();
  return {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: `peer-${revision}`,
    actor: "peer",
    baseRev: revision - 1,
    revision,
    ops: commit.ops.map((op) => captureDocumentOp(op)),
    effects: applied.effects,
  };
};

/** The operation kinds the peer's delivered batches held. */
type Tally = Map<string, number>;

const followsEveryPeerEdit = (
  tally: Tally,
  selections: readonly (readonly [number, number])[],
  steps: readonly Step[],
) => {
  const local = createCanonicalSession(fixture()).unwrap();
  const peer = createCanonicalSession(fixture()).unwrap();
  let fresh = 0x10000000;
  const freshId = () => {
    fresh += 1;
    return fresh.toString(16).toUpperCase();
  };
  // What the model expects of each story's selection, by story index.
  const expected: Ends[] = STORIES.map((story, index) => {
    const [anchor, head] = selections.at(index) ?? panic("No selection.");
    return {
      anchor: placeIn(local.document, story, anchor),
      head: placeIn(local.document, story, head),
    };
  });
  const expectedAt = (index: number): Ends =>
    expected.at(index) ?? panic("No expected selection.");
  let body = stateAt(local, OP_STORIES.MAIN, expectedAt(0));
  const record = (index: number) => {
    const story = STORIES.at(index) ?? panic("No story.");
    if (story === OP_STORIES.MAIN) {
      body = stateAt(local, story, expectedAt(index));
      return;
    }
    local
      .recordStorySelection(stateAt(local, story, expectedAt(index)), story)
      .unwrap();
  };
  for (const index of STORIES.keys()) record(index);
  let revision = 0;
  for (const step of steps) {
    const story = STORIES.at(step.story) ?? panic("No story.");
    if (step.type === "select") {
      expected[step.story] = {
        anchor: placeIn(local.document, story, step.anchor),
        head: placeIn(local.document, story, step.head),
      };
      record(step.story);
      continue;
    }
    const edit = peerEdit(peer.document, story, step, freshId);
    if (edit === undefined) continue;
    const before = peer.document;
    const batch = peerBatch(peer, story, edit, revision + 1);
    if (batch === undefined) continue;
    revision += 1;
    for (const op of batch.ops)
      tally.set(op.type, (tally.get(op.type) ?? 0) + 1);
    const commit = local
      .prepareRemote(body, { type: "rebased", batch })
      .unwrap();
    body = body.apply(commit.transaction);
    expect(commit.publish().isOk()).toBe(true);
    const ends = expectedAt(step.story);
    const actual = (index: number) =>
      endsOf(
        index === 0
          ? local.selection
          : local.storySelection(SECONDARY.at(index - 1) ?? panic("No story.")),
      ) ?? panic("The session holds no selection.");
    const shown = actual(step.story);
    for (const end of ["anchor", "head"] as const)
      expect(
        candidates(ends[end], edit, before, peer.document, story),
      ).toContainEqual(shown[end]);
    expected[step.story] = shown;
    // The body's selection is the editor's; the others are the session's.
    expect(endsOf(local.selection)).toEqual(expectedAt(0));
    for (const secondary of SECONDARY) {
      const index = STORIES.indexOf(secondary);
      const held = local.storySelection(secondary);
      expect(endsOf(held)).toEqual(expectedAt(index));
      // Each one still shows in its story's projection.
      const projection = local.projectStory(secondary).unwrap();
      expect(
        held !== null &&
          projection.positionAt(held.anchor).isOk() &&
          projection.positionAt(held.head).isOk(),
      ).toBe(true);
    }
  }
};

const law = (tally: Tally) =>
  fc.property(
    selectionsArbitrary,
    fc.array(stepArbitrary, { minLength: 1, maxLength: 16 }),
    (selections, steps) => {
      followsEveryPeerEdit(tally, selections, steps);
    },
  );

test("pinned: a peer types ahead of the caret in an open note and header", () => {
  const tally: Tally = new Map();
  followsEveryPeerEdit(
    tally,
    [
      [0, 0],
      // The end of "Header seed" and of "Footnote seed".
      [11 * 7, 11 * 7],
      [13 * 7, 13 * 7],
    ],
    [
      // Typed ahead of the note's own mark, as at the start of a saved note.
      { type: "peer", story: 2, kind: 0, seed: 0, text: "R" },
      { type: "peer", story: 2, kind: 0, seed: 7, text: "emote " },
      { type: "peer", story: 1, kind: 0, seed: 0, text: "Remote " },
    ],
  );
  expect(tally.get("insertContent") ?? 0).toBeGreaterThan(0);
});

test("on a fixed sample the peer inserts, deletes, splits and joins", () => {
  const tally: Tally = new Map();
  checkCoverageSample(law(tally), 30);
  // Typed text is inserted as inline content.
  const kinds = ["insertContent", "deleteRange", "splitBlock", "joinBlocks"];
  expect(kinds.filter((kind) => (tally.get(kind) ?? 0) === 0)).toEqual([]);
});

test(
  "every selection follows the text across a peer's edits in every story",
  async () => {
    await assertProperty(law(new Map()), {
      id: "every selection follows the text across a peer's edits in every story",
      numRuns: 60,
    });
  },
  propertyTestTimeout(60_000),
);
