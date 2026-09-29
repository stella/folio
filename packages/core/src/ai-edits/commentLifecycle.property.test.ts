/**
 * What the reviewer shows of every comment thread is what the saved package
 * reopens with, after every step of an editing session.
 *
 * `editKeepsCommentRanges.property` checks one edit against the saved markers.
 * Two defects got past it: its spans never crossed into a table, and it never
 * compared the live `getComments()` with the reopened one, so a paragraph a
 * range gained only on save and a thread whose anchor a rejection removed
 * (listed live, written back as a definition nothing references) both passed.
 * Nor did any step resolve tracked content a comment covered.
 *
 * Here the document holds a table between prose and a footnote with its own
 * comments; spans run across the table's edges; and a session of steps edits
 * body and note (direct and tracked), comments on and replies to blocks,
 * accepts and rejects single changes or everything, optionally reopening the
 * saved package between steps. After every step:
 *
 * - live and reopened `getComments()` agree on every thread: id, text,
 *   anchored text, the block it starts in, replies;
 * - the saved package has balanced ranges, and lists no thread that nothing
 *   in any story references (a ghost).
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type {
  BlockContent,
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
} from "../types/document";
import { FolioDocxReviewer, type FolioReviewComment } from "./headless";
import type { FolioAIEditOperation } from "./types";

setDefaultTimeout(propertyTestTimeout(30_000));

const FOOTNOTE_ID = 2;
const NOTE = { type: "footnote", noteId: FOOTNOTE_ID } as const;
/** Body paragraphs in document order; indices 2 and 3 are the table's cells. */
const BODY_SLOTS = 6;
const TABLE_SLOTS = new Set([2, 3]);
const NOTE_SLOTS = 3;
const EDIT_MODES = ["direct", "tracked-changes"] as const;
const EDIT_KINDS = ["replaceInBlock", "replaceBlock", "insertAfterBlock", "deleteBlock"] as const;

type Story = "body" | "note";
type GeneratedComment = { story: Story; first: number; last: number; text: string };
type Step =
  | {
      type: "edit";
      story: Story;
      kind: (typeof EDIT_KINDS)[number];
      mode: (typeof EDIT_MODES)[number];
      pick: number;
    }
  | { type: "comment"; story: Story; pick: number }
  | { type: "reply"; pick: number }
  | { type: "resolveAll"; how: "accept" | "reject" }
  | { type: "resolveOne"; how: "accept" | "reject"; pick: number };
type GeneratedSession = {
  comments: readonly GeneratedComment[];
  steps: readonly { step: Step; reopen: boolean }[];
};

/** A span; one paragraph half the time, since removing all it covers is the lifecycle's edge. */
const span = (slots: number) =>
  fc.oneof(
    fc.nat({ max: slots - 1 }).map((slot) => ({ first: slot, last: slot })),
    fc
      .tuple(fc.nat({ max: slots - 1 }), fc.nat({ max: slots - 1 }))
      .map(([one, other]) => ({ first: Math.min(one, other), last: Math.max(one, other) })),
  );

const commentArbitrary: fc.Arbitrary<GeneratedComment> = fc
  .oneof(
    {
      weight: 3,
      arbitrary: fc.record({ story: fc.constant<Story>("body"), span: span(BODY_SLOTS) }),
    },
    {
      weight: 1,
      arbitrary: fc.record({ story: fc.constant<Story>("note"), span: span(NOTE_SLOTS) }),
    },
  )
  .chain(({ story, span: { first, last } }) =>
    fc.stringMatching(/^[A-Za-z]{1,10}$/u).map((text) => ({ story, first, last, text })),
  );

const storyArbitrary = fc.oneof(
  { weight: 3, arbitrary: fc.constant<Story>("body") },
  { weight: 1, arbitrary: fc.constant<Story>("note") },
);
const pick = fc.nat({ max: 50 });
const how = fc.constantFrom("accept" as const, "reject" as const);

const stepArbitrary: fc.Arbitrary<Step> = fc.oneof(
  {
    weight: 4,
    arbitrary: fc.record({
      type: fc.constant("edit" as const),
      story: storyArbitrary,
      kind: fc.constantFrom(...EDIT_KINDS),
      mode: fc.constantFrom(...EDIT_MODES),
      pick,
    }),
  },
  {
    weight: 2,
    arbitrary: fc.record({ type: fc.constant("comment" as const), story: storyArbitrary, pick }),
  },
  { weight: 1, arbitrary: fc.record({ type: fc.constant("reply" as const), pick }) },
  { weight: 2, arbitrary: fc.record({ type: fc.constant("resolveAll" as const), how }) },
  { weight: 2, arbitrary: fc.record({ type: fc.constant("resolveOne" as const), how, pick }) },
);

const sessionArbitrary: fc.Arbitrary<GeneratedSession> = fc.record({
  comments: fc.array(commentArbitrary, { minLength: 1, maxLength: 4 }),
  steps: fc.array(fc.record({ step: stepArbitrary, reopen: fc.boolean() }), {
    minLength: 1,
    maxLength: 4,
  }),
});

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const slotText = (story: Story, index: number): string => {
  if (story === "note") {
    return `Note paragraph ${index}.`;
  }
  return TABLE_SLOTS.has(index)
    ? `Cell ${index} of the schedule.`
    : `Paragraph ${index} of the agreement.`;
};

/** Paragraphs for one story's slots, each carrying the markers its comments open or close. */
const storyParagraphs = (
  story: Story,
  slots: number,
  comments: readonly GeneratedComment[],
  paraIdPrefix: string,
): Paragraph[] =>
  Array.from({ length: slots }, (_, index) => {
    const content: ParagraphContent[] = [];
    for (const [id, comment] of comments.entries()) {
      if (comment.story === story && comment.first === index) {
        content.push({ type: "commentRangeStart", id });
      }
    }
    content.push(run(slotText(story, index)));
    if (story === "body" && index === BODY_SLOTS - 1) {
      content.push({ type: "run", content: [{ type: "footnoteRef", id: FOOTNOTE_ID }] });
    }
    for (const [id, comment] of comments.entries()) {
      if (comment.story === story && comment.last === index) {
        content.push({ type: "commentRangeEnd", id }, { type: "commentReference", id });
      }
    }
    return { type: "paragraph", paraId: `${paraIdPrefix}${index}`, content };
  });

const buildDocument = (comments: readonly GeneratedComment[]): Document => {
  const body = storyParagraphs("body", BODY_SLOTS, comments, "2200000");
  const [first, second, cellA, cellB, ...rest] = body;
  const content: BlockContent[] = [
    first!,
    second!,
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            { type: "tableCell", content: [cellA!] },
            { type: "tableCell", content: [cellB!] },
          ],
        },
      ],
    },
    ...rest,
  ];
  const authored: Comment[] = comments.map(({ text }, id) => ({
    id,
    author: "Dana Lindqvist",
    initials: "DL",
    date: "2024-01-01T00:00:00Z",
    content: [{ type: "paragraph", content: [run(text)] }],
  }));
  return {
    package: {
      document: { content, comments: authored },
      footnotes: [
        {
          type: "footnote",
          id: FOOTNOTE_ID,
          content: storyParagraphs("note", NOTE_SLOTS, comments, "2300000"),
        },
      ],
    },
  };
};

const blocksOf = (reviewer: FolioDocxReviewer, story: Story) =>
  (story === "body" ? reviewer.snapshot() : reviewer.snapshotStory(NOTE))?.blocks ?? [];

const editFor = (
  { kind }: Extract<Step, { type: "edit" }>,
  block: { id: string; text: string },
): FolioAIEditOperation | null => {
  switch (kind) {
    case "replaceInBlock":
      return block.text.length === 0
        ? null
        : {
            id: "edit",
            type: "replaceInBlock",
            blockId: block.id,
            find: block.text,
            replace: "Revised.",
          };
    case "replaceBlock":
      return { id: "edit", type: "replaceBlock", blockId: block.id, text: "Revised." };
    case "insertAfterBlock":
      return { id: "edit", type: "insertAfterBlock", blockId: block.id, text: "Inserted clause." };
    case "deleteBlock":
      return { id: "edit", type: "deleteBlock", blockId: block.id };
    default: {
      const unreachable: never = kind;
      throw new Error(`Unhandled edit kind: ${String(unreachable)}`);
    }
  }
};

const applyStep = (reviewer: FolioDocxReviewer, step: Step): void => {
  switch (step.type) {
    case "edit":
    case "comment": {
      const blocks = blocksOf(reviewer, step.story);
      const block = blocks[step.pick % Math.max(blocks.length, 1)];
      if (!block) {
        return;
      }
      const operation =
        step.type === "edit"
          ? editFor(step, block)
          : ({
              id: "comment",
              type: "commentOnBlock",
              blockId: block.id,
              comment: { text: "Added in session." },
            } satisfies FolioAIEditOperation);
      if (!operation) {
        return;
      }
      const mode = step.type === "edit" ? step.mode : "direct";
      if (step.story === "body") {
        reviewer.applyOperations([operation], { mode });
      } else {
        reviewer.applyDocumentOperationsToStory({
          story: NOTE,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: [operation],
          },
        });
      }
      return;
    }
    case "reply": {
      const threads = reviewer.getComments();
      const target = threads[step.pick % Math.max(threads.length, 1)];
      if (target) {
        reviewer.replyTo(target, { text: "Session reply." });
      }
      return;
    }
    case "resolveAll":
      if (step.how === "accept") {
        reviewer.acceptAll();
      } else {
        reviewer.rejectAll();
      }
      return;
    case "resolveOne": {
      const changes = reviewer.getChanges();
      const change = changes[step.pick % Math.max(changes.length, 1)];
      if (!change) {
        return;
      }
      if (step.how === "accept") {
        reviewer.acceptChange(change);
      } else {
        reviewer.rejectChange(change);
      }
      return;
    }
    default: {
      const unreachable: never = step;
      throw new Error(`Unhandled step: ${JSON.stringify(unreachable)}`);
    }
  }
};

/** The thread fields the reviewer and the package must agree on. */
const threadsOf = (comments: readonly FolioReviewComment[]) =>
  comments.map(({ id, author, text, anchoredText, blockId, replies, done }) => ({
    id,
    blockId,
    author,
    text,
    anchoredText,
    done,
    replies: replies.map((reply) => ({ id: reply.id, text: reply.text })),
  }));

const ids = (xml: string, element: string): number[] =>
  [...xml.matchAll(new RegExp(`<w:${element}\\s[^>]*w:id="(-?\\d+)"`, "gu"))].map(([, id]) =>
    Number(id),
  );

/** Every range opens before it closes, in each story part on its own. */
const expectBalanced = (part: string, xml: string): void => {
  const starts = ids(xml, "commentRangeStart");
  const ends = ids(xml, "commentRangeEnd");
  for (const id of new Set([...starts, ...ends])) {
    const start = xml.search(new RegExp(`<w:commentRangeStart\\s[^>]*w:id="${id}"`, "u"));
    const end = xml.search(new RegExp(`<w:commentRangeEnd\\s[^>]*w:id="${id}"`, "u"));
    expect({ part, id, opens: start >= 0, closes: end >= 0, ordered: start < end }).toEqual({
      part,
      id,
      opens: true,
      closes: true,
      ordered: true,
    });
  }
};

/**
 * Balanced ranges in every story part, and no thread the package lists that
 * nothing in any story references: a ghost a reader cannot scroll to. Replies
 * are threaded by `commentsExtended.xml`, so only thread roots are checked.
 */
const expectPackageConsistent = async (
  saved: ArrayBuffer,
  threads: readonly { id: number }[],
): Promise<void> => {
  const zip = await JSZip.loadAsync(saved);
  const read = async (path: string) => (await zip.file(path)?.async("text")) ?? "";
  const documentXml = await read("word/document.xml");
  const footnotesXml = await read("word/footnotes.xml");
  expectBalanced("document", documentXml);
  expectBalanced("footnotes", footnotesXml);
  const referenced = new Set<number>();
  for (const xml of [documentXml, footnotesXml]) {
    for (const id of [...ids(xml, "commentRangeStart"), ...ids(xml, "commentReference")]) {
      referenced.add(id);
    }
  }
  const ghosts = threads.map(({ id }) => id).filter((id) => !referenced.has(id));
  expect({ ghosts }).toEqual({ ghosts: [] });
};

describe("live comment threads match the saved package after every step", () => {
  test("over generated sessions across a table and a note", async () => {
    await assertProperty(
      fc.asyncProperty(sessionArbitrary, async ({ comments, steps }) => {
        let reviewer = await FolioDocxReviewer.fromBuffer(
          await createDocx(buildDocument(comments)),
          { author: "Editor" },
        );
        for (const [index, { step, reopen }] of steps.entries()) {
          applyStep(reviewer, step);
          const live = threadsOf(reviewer.getComments());
          const saved = await reviewer.toBuffer();
          const reopened = await FolioDocxReviewer.fromBuffer(saved, { author: "Editor" });
          const persisted = threadsOf(reopened.getComments());
          await expectPackageConsistent(saved, persisted);
          expect({ step: index, threads: persisted }).toEqual({ step: index, threads: live });
          if (reopen) {
            reviewer = reopened;
          }
        }
      }),
      { numRuns: 40 },
    );
  }, 240_000);
});
