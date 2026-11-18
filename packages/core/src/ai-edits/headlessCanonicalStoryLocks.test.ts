/** Every canonical story uses the source lock owner and reports complete review outcomes. */
import { afterEach, expect, setSystemTime, test } from "bun:test";

import { CanonicalSessionError } from "../canonical/canonicalSession";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { CANONICAL_RULE } from "../types/canonicalCapabilities";
import type { BlockContent } from "../types/document";
import { FolioDocxReviewer } from "./headless";

afterEach(() => setSystemTime());

const STORIES = [
  { type: "main" },
  { type: "header", relationshipId: "rIdHeader" },
  { type: "footer", relationshipId: "rIdFooter" },
  { type: "footnote", noteId: 7 },
  { type: "endnote", noteId: 8 },
] as const;
type StoryKind = (typeof STORIES)[number]["type"];
const DATE = "2026-01-02T03:04:05Z";
const info = (id: number) => ({ id, author: "Writer", date: DATE });

type FixtureOptions = { locked: StoryKind; unlocked: "everyStory" | "none" };
const fixture = async ({ locked, unlocked }: FixtureOptions) => {
  const content = (kind: StoryKind): BlockContent[] => {
    const index = STORIES.findIndex(({ type }) => type === kind);
    const insertion = index * 10 + 101;
    const deletion = insertion + 1;
    return [
      paragraph(
        kind === locked
          ? [
              {
                type: "inlineSdt",
                properties: {
                  sdtType: "richText",
                  id: index + 11,
                  lock: "contentLocked",
                },
                content: [
                  run("Held "),
                  {
                    type: "insertion",
                    info: info(insertion),
                    content: [run("added")],
                  },
                  {
                    type: "deletion",
                    info: info(deletion),
                    content: [run("removed")],
                  },
                ],
              },
            ]
          : [run("Plain words.")],
      ),
      paragraph(
        unlocked === "everyStory"
          ? [
              run("Outside "),
              {
                type: "insertion",
                info: info(insertion + 2),
                content: [run("addition")],
              },
            ]
          : [run("Outside words.")],
      ),
    ];
  };
  const document = fromMarkdown("Body.");
  document.package.document.content = content("main");
  document.package.document.content.push(
    paragraph([
      {
        type: "run",
        content: [
          { type: "footnoteRef", id: 7 },
          { type: "endnoteRef", id: 8 },
        ],
      },
    ]),
  );
  document.package.document.finalSectionProperties = {
    headerReferences: [{ type: "default", rId: "rIdHeader" }],
    footerReferences: [{ type: "default", rId: "rIdFooter" }],
  };
  document.package.headers = new Map([
    [
      "rIdHeader",
      { type: "header", hdrFtrType: "default", content: content("header") },
    ],
  ]);
  document.package.footers = new Map([
    [
      "rIdFooter",
      { type: "footer", hdrFtrType: "default", content: content("footer") },
    ],
  ]);
  document.package.footnotes = [
    { type: "footnote", id: 7, content: content("footnote") },
  ];
  document.package.endnotes = [
    { type: "endnote", id: 8, content: content("endnote") },
  ];
  const prepared = await prepareCanonicalDocxInput(await createDocx(document));
  if (prepared.isErr()) throw prepared.error;
  return FolioDocxReviewer.fromBuffer(prepared.value.slice().buffer, {
    session: "canonical",
  });
};

const changes = (
  reviewer: FolioDocxReviewer,
  story: (typeof STORIES)[number],
) =>
  reviewer.readReviewedStory({ story, view: "current-markup" })?.changes ?? [];
const reads = (reviewer: FolioDocxReviewer) => ({
  document: reviewer.toDocument(),
  stories: reviewer.listStories(),
  changes: STORIES.map((story) => changes(reviewer, story)),
  suggestions: reviewer.exportPendingSuggestions(),
});
const expectLockRefusal = (action: () => unknown) => {
  let thrown: unknown;
  try {
    action();
  } catch (error) {
    thrown = error;
  }
  expect(CanonicalSessionError.is(thrown)).toBe(true);
  expect(thrown).toMatchObject({
    reason: "refused",
    ground: { rule: CANONICAL_RULE.contentControlLock },
  });
};

for (const [index, story] of STORIES.entries())
  for (const decision of ["accept", "reject"] as const) {
    test(`${story.type} ${decision}: bulk review resolves other changes and reports every held change`, async () => {
      const reviewer = await fixture({
        locked: story.type,
        unlocked: "everyStory",
      });
      expect(reviewer.listStories().map(({ handle }) => handle)).toEqual(
        STORIES,
      );
      const held = changes(reviewer, story).filter(
        ({ id }) => id !== index * 10 + 103,
      );
      expect(held).toHaveLength(2);
      const result =
        decision === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
      expect(result.resolved).toBe(STORIES.length);
      expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual(
        held.map(({ id }) => ({ id, reason: "contentControlLock" })),
      );
      for (const other of STORIES)
        expect(changes(reviewer, other)).toEqual(other === story ? held : []);
    });

    test(`${story.type} ${decision}: a package containing only held changes reports them explicitly`, async () => {
      const reviewer = await fixture({ locked: story.type, unlocked: "none" });
      const before = reads(reviewer);
      const held = changes(reviewer, story);
      expect(held).toHaveLength(2);
      const result =
        decision === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
      expect(result.resolved).toBe(0);
      expect(result.skipped.map(({ id }) => id)).toEqual(
        held.map(({ id }) => id),
      );
      expect(reads(reviewer)).toEqual(before);
    });
  }

for (const story of STORIES)
  for (const view of ["original", "final"] as const) {
    test(`${story.type} ${view}: an incomplete full view refuses without live mutation or history changes`, async () => {
      setSystemTime(new Date(DATE));
      const reviewer = await fixture({
        locked: story.type,
        unlocked: "everyStory",
      });
      const beforeEdit = reads(reviewer);
      const anchor = reviewer
        .getContent()
        .find((block) => block.text.startsWith("Outside"));
      if (!anchor) throw new Error("The fixture has no unlocked anchor");
      const edited = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: [
            {
              id: "history",
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "Independent edit.",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(edited.applied).toHaveLength(1);
      if (edited.undoHandle === null)
        throw new Error("The edit has no undo handle");
      const before = reads(reviewer);
      const saved = new Uint8Array(await reviewer.toBuffer());
      expectLockRefusal(() => reviewer.resolveReviewedStory({ story, view }));
      expect(reads(reviewer)).toEqual(before);
      expectLockRefusal(() => reviewer.readReviewedStory({ story, view }));
      expectLockRefusal(() => reviewer.storyTables({ story, view }));
      expectLockRefusal(() =>
        reviewer.toResolvedDocument(view === "final" ? "accept" : "reject"),
      );
      expect(reads(reviewer)).toEqual(before);
      expect(new Uint8Array(await reviewer.toBuffer())).toEqual(saved);
      expect(reviewer.undoDocumentOperations(edited.undoHandle)).toMatchObject({
        status: "undone",
      });
      expect(reads(reviewer)).toEqual(beforeEdit);

      const other = STORIES.find((candidate) => candidate.type !== story.type);
      if (!other) throw new Error("The fixture has no other story");
      const held = changes(reviewer, story);
      expect(reviewer.resolveReviewedStory({ story: other, view })).toBe(true);
      expect(changes(reviewer, other)).toEqual([]);
      expect(changes(reviewer, story)).toEqual(held);
    });
  }
