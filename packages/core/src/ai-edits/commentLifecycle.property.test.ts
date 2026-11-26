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
 *   in any story references (a ghost);
 * - accepting or rejecting everything still projects: no comment's marks
 *   disagree with its explicit boundary nodes.
 *
 * Every other paragraph ends in a link a range may start or end inside, and a
 * batch mixes paragraph, row and table structure edits, so a rejected split
 * can join a range's boundary into the link its other half left. A comment
 * added on any block (body, cell, inserted or beside a retracted one) anchors
 * the block's text.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
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
const EDIT_KINDS = [
  "replaceInBlock",
  "replaceBlock",
  "insertAfterBlock",
  "deleteBlock",
  "splitBlock",
  "mergeBlockWithNext",
  "insertTableRow",
  "deleteTableRow",
  "insertTable",
  "deleteTable",
] as const;
/** Slots whose paragraph ends in a link: body prose, a cell, a note paragraph. */
const LINK_SLOTS = new Set([1, 3, 5]);

type EditKind = (typeof EDIT_KINDS)[number];
type Story = "body" | "note";
type GeneratedComment = {
  story: Story;
  first: number;
  last: number;
  text: string;
  /** Whether each boundary sits inside its paragraph's link, when it has one. */
  startInLink: boolean;
  endInLink: boolean;
};
type Step =
  | {
      type: "edit";
      story: Story;
      mode: (typeof EDIT_MODES)[number];
      edits: readonly { kind: EditKind; pick: number; at: number }[];
    }
  | {
      type: "comment";
      story: Story;
      mode: (typeof EDIT_MODES)[number];
      pick: number;
    }
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
      .map(([one, other]) => ({
        first: Math.min(one, other),
        last: Math.max(one, other),
      })),
  );

/**
 * `structure` sessions place boundaries inside links and batch structure
 * edits; without it a session draws exactly what the pinned seeds replay.
 */
const commentArbitrary = (structure: boolean): fc.Arbitrary<GeneratedComment> =>
  fc
    .oneof(
      {
        weight: 3,
        arbitrary: fc.record({
          story: fc.constant<Story>("body"),
          span: span(BODY_SLOTS),
        }),
      },
      {
        weight: 1,
        arbitrary: fc.record({
          story: fc.constant<Story>("note"),
          span: span(NOTE_SLOTS),
        }),
      },
    )
    .chain(({ story, span: { first, last } }) => {
      const text = fc.stringMatching(/^[A-Za-z]{1,10}$/u);
      return structure
        ? fc
            .record({
              text,
              startInLink: fc.boolean(),
              endInLink: fc.boolean(),
            })
            .map(({ text: value, startInLink, endInLink }) => ({
              story,
              first,
              last,
              text: value,
              startInLink,
              endInLink,
            }))
        : text.map((value) => ({
            story,
            first,
            last,
            text: value,
            startInLink: false,
            endInLink: false,
          }));
    });

const storyArbitrary = fc.oneof(
  { weight: 3, arbitrary: fc.constant<Story>("body") },
  { weight: 1, arbitrary: fc.constant<Story>("note") },
);
const pick = fc.nat({ max: 50 });
const how = fc.constantFrom("accept" as const, "reject" as const);
const modeArbitrary = fc.constantFrom(...EDIT_MODES);

const editArbitrary = (structure: boolean): fc.Arbitrary<Step> =>
  structure
    ? fc.record({
        type: fc.constant("edit" as const),
        story: storyArbitrary,
        mode: modeArbitrary,
        edits: fc.array(
          fc.record({ kind: fc.constantFrom(...EDIT_KINDS), pick, at: pick }),
          { minLength: 1, maxLength: 3 },
        ),
      })
    : fc
        .record({
          type: fc.constant("edit" as const),
          story: storyArbitrary,
          kind: fc.constantFrom(...EDIT_KINDS.slice(0, 4)),
          mode: modeArbitrary,
          pick,
        })
        .map(({ type, story, kind, mode, pick: at }) => ({
          type,
          story,
          mode,
          edits: [{ kind, pick: at, at: 0 }],
        }));

const commentStepArbitrary = (structure: boolean): fc.Arbitrary<Step> =>
  structure
    ? fc.record({
        type: fc.constant("comment" as const),
        story: storyArbitrary,
        mode: modeArbitrary,
        pick,
      })
    : fc
        .record({
          type: fc.constant("comment" as const),
          story: storyArbitrary,
          pick,
        })
        .map(({ type, story, pick: at }) => ({
          type,
          story,
          mode: "direct" as const,
          pick: at,
        }));

const stepArbitrary = (structure: boolean): fc.Arbitrary<Step> =>
  fc.oneof(
    { weight: 4, arbitrary: editArbitrary(structure) },
    { weight: 2, arbitrary: commentStepArbitrary(structure) },
    {
      weight: 1,
      arbitrary: fc.record({ type: fc.constant("reply" as const), pick }),
    },
    {
      weight: 2,
      arbitrary: fc.record({ type: fc.constant("resolveAll" as const), how }),
    },
    {
      weight: 2,
      arbitrary: fc.record({
        type: fc.constant("resolveOne" as const),
        how,
        pick,
      }),
    },
  );

const sessionArbitrary = (structure: boolean): fc.Arbitrary<GeneratedSession> =>
  fc.record({
    comments: fc.array(commentArbitrary(structure), {
      minLength: 1,
      maxLength: 4,
    }),
    steps: fc.array(
      fc.record({ step: stepArbitrary(structure), reopen: fc.boolean() }),
      { minLength: 1, maxLength: 4 },
    ),
  });

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const slotText = (story: Story, index: number): string => {
  if (story === "note") {
    return `Note paragraph ${index} of the note.`;
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
    const ours = [...comments.entries()].filter(
      ([, comment]) => comment.story === story,
    );
    const linked = LINK_SLOTS.has(index);
    const starts = (inLink: boolean) =>
      ours.flatMap(([id, comment]) =>
        comment.first === index && (linked && comment.startInLink) === inLink
          ? [{ type: "commentRangeStart" as const, id }]
          : [],
      );
    const ends = (inLink: boolean) =>
      ours.flatMap(([id, comment]) =>
        comment.last === index && (linked && comment.endInLink) === inLink
          ? [{ type: "commentRangeEnd" as const, id }]
          : [],
      );
    content.push(...starts(false));
    const [head, tail] = slotText(story, index).split(" of ");
    if (linked)
      content.push(run(`${head} `), {
        type: "hyperlink",
        href: "https://example.com/terms",
        children: [...starts(true), run(`of ${tail}`), ...ends(true)],
      });
    else content.push(run(slotText(story, index)));
    if (story === "body" && index === BODY_SLOTS - 1) {
      content.push({
        type: "run",
        content: [{ type: "footnoteRef", id: FOOTNOTE_ID }],
      });
    }
    content.push(...ends(false));
    for (const [id, comment] of ours) {
      if (comment.last === index) {
        content.push({ type: "commentReference", id });
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
  (story === "body" ? reviewer.snapshot() : reviewer.snapshotStory(NOTE))
    ?.blocks ?? [];

const editFor = (
  { kind, at }: { kind: EditKind; at: number },
  block: { id: string; text: string },
  id: string,
): FolioAIEditOperation | null => {
  const blockId = block.id;
  switch (kind) {
    case "replaceInBlock":
      return block.text.length === 0
        ? null
        : { id, type: kind, blockId, find: block.text, replace: "Revised." };
    case "replaceBlock":
      return { id, type: kind, blockId, text: "Revised." };
    case "insertAfterBlock":
      return { id, type: kind, blockId, text: "Inserted clause." };
    case "splitBlock":
      return block.text.length < 2
        ? null
        : {
            id,
            type: kind,
            blockId,
            offset: 1 + (at % (block.text.length - 1)),
          };
    case "insertTableRow":
      return {
        id,
        type: kind,
        blockId,
        position: at % 2 === 0 ? "before" : "after",
        cellTexts: ["Row cell one", "Row cell two"],
      };
    case "insertTable":
      return { id, type: kind, blockId, rows: [["Term", "Value"]] };
    case "deleteBlock":
    case "mergeBlockWithNext":
    case "deleteTableRow":
    case "deleteTable":
      return { id, type: kind, blockId };
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
      const blockAt = (at: number) => blocks[at % Math.max(blocks.length, 1)];
      const block = blockAt(step.type === "edit" ? 0 : step.pick);
      if (!block) {
        return;
      }
      const operations =
        step.type === "edit"
          ? step.edits.flatMap((edit, index) => {
              const target = blockAt(edit.pick);
              const operation =
                target && editFor(edit, target, `edit-${index}`);
              return operation ? [operation] : [];
            })
          : [
              {
                id: "comment",
                type: "commentOnBlock",
                blockId: block.id,
                comment: { text: "Added in session." },
              } satisfies FolioAIEditOperation,
            ];
      const before = new Set(reviewer.getComments().map(({ id }) => id));
      const { mode } = step;
      const applied =
        step.story === "body"
          ? reviewer.applyOperations(operations, {
              undefinedReferences: "refuse",
              mode,
            }).applied
          : reviewer.applyDocumentOperationsToStory({
              undefinedReferences: "refuse",
              story: NOTE,
              batch: {
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                mode,
                operations,
              },
            }).applied;
      // A comment on a block anchors that block's text, wherever it sits.
      if (step.type === "comment" && applied.length > 0) {
        const added = reviewer
          .getComments()
          .filter(({ id }) => !before.has(id))
          .map(({ blockId, anchoredText }) => ({ blockId, anchoredText }));
        expect(added).toEqual([
          { blockId: block.id, anchoredText: block.text },
        ]);
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
  comments.map(
    ({ id, author, text, anchoredText, blockId, replies, done }) => ({
      id,
      blockId,
      author,
      text,
      anchoredText,
      done,
      replies: replies.map((reply) => ({ id: reply.id, text: reply.text })),
    }),
  );

const ids = (xml: string, element: string): number[] =>
  [
    ...xml.matchAll(new RegExp(`<w:${element}\\s[^>]*w:id="(-?\\d+)"`, "gu")),
  ].map(([, id]) => Number(id));

/** Every range opens before it closes, in each story part on its own. */
const expectBalanced = (part: string, xml: string): void => {
  const starts = ids(xml, "commentRangeStart");
  const ends = ids(xml, "commentRangeEnd");
  for (const id of new Set([...starts, ...ends])) {
    const start = xml.search(
      new RegExp(`<w:commentRangeStart\\s[^>]*w:id="${id}"`, "u"),
    );
    const end = xml.search(
      new RegExp(`<w:commentRangeEnd\\s[^>]*w:id="${id}"`, "u"),
    );
    expect({
      part,
      id,
      opens: start >= 0,
      closes: end >= 0,
      ordered: start < end,
    }).toEqual({
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
  const zip = await ZipArchive.read(saved);
  const read = async (path: string) => (await zip.entry(path)?.text()) ?? "";
  const documentXml = await read("word/document.xml");
  const footnotesXml = await read("word/footnotes.xml");
  expectBalanced("document", documentXml);
  expectBalanced("footnotes", footnotesXml);
  const referenced = new Set<number>();
  for (const xml of [documentXml, footnotesXml]) {
    for (const id of [
      ...ids(xml, "commentRangeStart"),
      ...ids(xml, "commentReference"),
    ]) {
      referenced.add(id);
    }
  }
  const ghosts = threads
    .map(({ id }) => id)
    .filter((id) => !referenced.has(id));
  expect({ ghosts }).toEqual({ ghosts: [] });
};

describe("live comment threads match the saved package after every step", () => {
  test("deleting a replied-to paragraph retires its entire thread", async () => {
    const plain = { startInLink: false, endInLink: false };
    const comments: GeneratedComment[] = [
      { story: "body", first: 1, last: 1, text: "A", ...plain },
      { story: "body", first: 0, last: 0, text: "A", ...plain },
    ];
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(buildDocument(comments)),
      { author: "Editor" },
    );

    const expectSavedRoots = async (expectedRootIds: number[]) => {
      const saved = await reviewer.toBuffer();
      const reopened = await FolioDocxReviewer.fromBuffer(saved, {
        author: "Editor",
      });
      const persisted = reopened.toDocument().package.document.comments ?? [];
      const persistedIds = new Set(persisted.map(({ id }) => id));
      expect(
        persisted.every(
          ({ parentId }) =>
            parentId === undefined || persistedIds.has(parentId),
        ),
      ).toBe(true);
      expect(reopened.getComments().map(({ id }) => id)).toEqual(
        expectedRootIds,
      );
      await expectPackageConsistent(saved, reopened.getComments());
    };

    expect(reviewer.replyTo(0, { text: "Session reply." })).not.toBeNull();
    await expectSavedRoots([0, 1]);
    applyStep(reviewer, {
      type: "edit",
      story: "body",
      mode: "direct",
      edits: [{ kind: "deleteBlock", pick: 7, at: 0 }],
    });
    await expectSavedRoots([1]);
    applyStep(reviewer, { type: "resolveAll", how: "accept" });
    await expectSavedRoots([1]);
  });

  const checkSession = async ({ comments, steps }: GeneratedSession) => {
    let reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(buildDocument(comments)),
      { author: "Editor" },
    );
    for (const [index, { step, reopen }] of steps.entries()) {
      applyStep(reviewer, step);
      for (const decision of ["accept", "reject"] as const)
        reviewer.toResolvedDocument(decision);
      const live = threadsOf(reviewer.getComments());
      const saved = await reviewer.toBuffer();
      const reopened = await FolioDocxReviewer.fromBuffer(saved, {
        author: "Editor",
      });
      const persisted = threadsOf(reopened.getComments());
      await expectPackageConsistent(saved, persisted);
      expect({ step: index, threads: persisted }).toEqual({
        step: index,
        threads: live,
      });
      if (reopen) {
        reviewer = reopened;
      }
    }
  };

  const inLink = (
    first: number,
    last: number,
    startInLink: boolean,
    endInLink: boolean,
  ): GeneratedComment => ({
    story: "body",
    first,
    last,
    text: "A",
    startInLink,
    endInLink,
  });

  test("a range with one boundary inside a link opens and saves", async () => {
    await checkSession({ comments: [inLink(1, 1, true, false)], steps: [] });
    await checkSession({ comments: [inLink(1, 1, false, true)], steps: [] });
  });

  test("a range runs across a paragraph whose own range sits in its link", async () => {
    await checkSession({
      comments: [inLink(0, 4, false, false), inLink(1, 1, true, true)],
      steps: [],
    });
  });

  test("deleting a table keeps the reference of a range that ends in it", async () => {
    // The first range ends in the table's first cell, where the second
    // starts: both boundaries stay, and so does the first's reference.
    await checkSession({
      comments: [inLink(0, 2, false, false), inLink(2, 4, false, false)],
      steps: [
        {
          step: {
            type: "edit",
            story: "body",
            mode: "direct",
            edits: [{ kind: "deleteTable", pick: 2, at: 0 }],
          },
          reopen: false,
        },
      ],
    });
  });

  test("rejecting a split joins a comment's start into the link it divided", async () => {
    // Split slot 1 inside its link, comment on the tail (it opens before the
    // tail's link), then reject: the start lands inside the joined link.
    await checkSession({
      comments: [inLink(0, 0, false, false)],
      steps: [
        {
          step: {
            type: "edit",
            story: "body",
            mode: "tracked-changes",
            edits: [{ kind: "splitBlock", pick: 1, at: 14 }],
          },
          reopen: false,
        },
        {
          step: {
            type: "comment",
            story: "body",
            mode: "tracked-changes",
            pick: 2,
          },
          reopen: false,
        },
        { step: { type: "resolveAll", how: "reject" }, reopen: false },
      ],
    });
  });

  test(
    "over generated sessions across a table and a note",
    async () => {
      await assertProperty(
        fc.asyncProperty(sessionArbitrary(false), checkSession),
        {
          numRuns: 40,
          id: "over generated sessions across a table and a note",
        },
      );
    },
    propertyTestTimeout(240_000),
  );

  test(
    "over structure batches beside ranges that start or end inside links",
    async () => {
      await assertProperty(
        fc.asyncProperty(sessionArbitrary(true), checkSession),
        {
          numRuns: 40,
          id: "over structure batches beside ranges that start or end inside links",
        },
      );
    },
    propertyTestTimeout(240_000),
  );
});
