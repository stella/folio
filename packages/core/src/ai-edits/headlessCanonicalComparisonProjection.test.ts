import { afterEach, expect, mock, setSystemTime, spyOn, test } from "bun:test";

import { applyDocumentOps } from "@stll/docx-core/ops";

import * as canonical from "../canonical/canonicalSession";
import type {
  CanonicalCommittedBatch,
  CanonicalSession,
} from "../canonical/canonicalSession";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { CANONICAL_GAP, CANONICAL_RULE } from "../types/canonicalCapabilities";
import type { BlockContent, Insertion } from "../types/document";
import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "./headless";

afterEach(() => {
  mock.restore();
  setSystemTime();
});
const DATE = "2026-01-02T03:04:05Z";
const STORIES = [
  { type: "main" },
  { type: "header", relationshipId: "rIdHeader" },
  { type: "footer", relationshipId: "rIdFooter" },
  { type: "footnote", noteId: 7 },
  { type: "endnote", noteId: 8 },
] as const;

type Observation = {
  session: CanonicalSession;
  batches: CanonicalCommittedBatch[];
  disposals: (() => number)[];
};
/** Observe real session construction and listeners; every original implementation still runs. */
const observeSessions = () => {
  const sessions: Observation[] = [];
  const create = canonical.createCanonicalSession;
  spyOn(canonical, "createCanonicalSession").mockImplementation((...args) => {
    const result = create(...args);
    if (result.isErr()) return result;
    const session = result.value;
    const observation: Observation = { session, batches: [], disposals: [] };
    sessions.push(observation);
    const subscribe = session.onCommitted.bind(session);
    spyOn(session, "onCommitted").mockImplementation((listener) => {
      const off = subscribe((batch) => {
        observation.batches.push(batch);
        listener(batch);
      });
      const dispose = mock(() => off());
      observation.disposals.push(() => dispose.mock.calls.length);
      return dispose;
    });
    return result;
  });
  return sessions;
};

const fixture = async (mode: "unlocked" | "lateLocked") => {
  const content = (kind: (typeof STORIES)[number]["type"]): BlockContent[] => {
    const index = STORIES.findIndex(({ type }) => type === kind);
    const insertion = {
      type: "insertion",
      info: { id: 100 + index, author: "Writer", date: DATE },
      content: [run(`${kind} addition`)],
    } satisfies Insertion;
    return [
      paragraph(
        mode === "lateLocked" && kind === "endnote"
          ? [
              {
                type: "inlineSdt",
                properties: {
                  sdtType: "richText",
                  id: 11,
                  lock: "contentLocked",
                },
                content: [insertion],
              },
            ]
          : [insertion],
      ),
      paragraph([run(`${kind} outside`)]),
    ];
  };
  const document = fromMarkdown("Body");
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
    pageWidth: 12240,
    pageHeight: 15840,
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
  const bytes = await prepareCanonicalDocxInput(await createDocx(document));
  if (bytes.isErr()) throw bytes.error;
  return FolioDocxReviewer.fromBuffer(bytes.value.slice().buffer, {
    session: "canonical",
    author: "Reviewer",
  });
};

const reads = (reviewer: FolioDocxReviewer) => ({
  document: reviewer.toDocument(),
  stories: reviewer.listStories(),
  pending: reviewer.exportPendingSuggestions(),
  reviewed: STORIES.map((story) =>
    reviewer.readReviewedStory({ story, view: "current-markup" }),
  ),
});
const firstSession = (observations: readonly Observation[]) => {
  const first = observations.at(0);
  if (!first)
    throw new TypeError("The real reviewer created its canonical session");
  return first;
};
const expectRefusal = (
  action: () => unknown,
  ground:
    | { rule: typeof CANONICAL_RULE.contentControlLock }
    | { gap: typeof CANONICAL_GAP.publicHeadlessSession },
) => {
  let thrown: unknown;
  try {
    action();
  } catch (error) {
    thrown = error;
  }
  expect(canonical.CanonicalSessionError.is(thrown)).toBe(true);
  expect(thrown).toMatchObject({ reason: "refused", ground });
};

test("comparison projection accounts for every story and publishes one reversible commit from an isolated fork", async () => {
  setSystemTime(new Date(DATE));
  const observations = observeSessions();
  const reviewer = await fixture("unlocked");
  const original = firstSession(observations);
  const before = cloneDocumentWithParagraphPropertySources(
    original.session.document,
  );
  const emitted: CanonicalCommittedBatch[] = [];
  const off = original.session.onCommitted((batch) => emitted.push(batch));
  const projected = getFolioDocxComparisonAccess(reviewer).projectStories(
    "with-revision-census",
  );
  off();
  expect(
    [...new Set(projected.stories.map(({ handle }) => handle.type))].sort(),
  ).toEqual(STORIES.map(({ type }) => type).sort());
  expect([...projected.revisions.ids].sort((a, b) => a - b)).toEqual([
    100, 101, 102, 103, 104,
  ]);
  expect(projected.revisions.present).toBe(true);
  expect(projected.stories.every(({ snapshot }) => snapshot !== null)).toBe(
    true,
  );
  expect(emitted).toHaveLength(1);
  expect(original.session.version).toBe(1);
  expect(original.session.canUndo).toBe(true);
  const batch = emitted.at(0);
  if (!batch) throw new TypeError("Projection published its one actual batch");
  const after = cloneDocumentWithParagraphPropertySources(
    original.session.document,
  );
  expect(applyDocumentOps(before, batch.ops).unwrap().document).toStrictEqual(
    after,
  );
  expect(
    applyDocumentOps(after, batch.inverse).unwrap().document,
  ).toStrictEqual(before);
  expect(
    applyDocumentOps(after, JSON.parse(JSON.stringify(batch.inverse))).unwrap()
      .document,
  ).toStrictEqual(before);
  for (const story of STORIES)
    expect(
      reviewer.readReviewedStory({ story, view: "current-markup" })?.changes,
    ).toEqual([]);
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    { session: "canonical" },
  );
  for (const story of STORIES) {
    expect(reopened.readStory(story)).toEqual(reviewer.readStory(story));
    expect(
      reopened.readReviewedStory({ story, view: "current-markup" })?.changes,
    ).toEqual([]);
  }
  const projectedFork = observations.at(1);
  if (!projectedFork)
    throw new TypeError("Projection created its independent fork");
  expect(projectedFork.batches.length).toBeGreaterThan(1);
  expect(projectedFork.disposals.map((count) => count())).toEqual([1]);
});

test("a late locked story refuses projection without publishing any earlier preflight resolution", async () => {
  setSystemTime(new Date(DATE));
  const observations = observeSessions();
  const reviewer = await fixture("lateLocked");
  const initial = reads(reviewer);
  const block = reviewer
    .getContent()
    .find(({ text }) => text === "main outside");
  if (!block)
    throw new TypeError("The source has its real unlocked main block");
  const edit = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        {
          id: "history",
          type: "insertAfterBlock",
          blockId: block.id,
          text: "Independent history.",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(edit.applied).toHaveLength(1);
  if (edit.undoHandle === null)
    throw new TypeError("The real edit produced its undo handle");
  const original = firstSession(observations);
  const version = original.session.version;
  const history = {
    undo: original.session.canUndo,
    redo: original.session.canRedo,
  };
  const before = reads(reviewer);
  const bytes = new Uint8Array(await reviewer.toBuffer());
  expectRefusal(
    () =>
      getFolioDocxComparisonAccess(reviewer).projectStories(
        "with-revision-census",
      ),
    { rule: CANONICAL_RULE.contentControlLock },
  );
  expect(reads(reviewer)).toStrictEqual(before);
  expect(original.session.version).toBe(version);
  expect({
    undo: original.session.canUndo,
    redo: original.session.canRedo,
  }).toEqual(history);
  expect(new Uint8Array(await reviewer.toBuffer())).toEqual(bytes);
  const projectedFork = observations.at(1);
  if (!projectedFork) throw new TypeError("The late refusal ran on its fork");
  expect(projectedFork.batches.length).toBeGreaterThan(1);
  expect(projectedFork.disposals.map((count) => count())).toEqual([1]);
  expect(reviewer.undoDocumentOperations(edit.undoHandle)).toMatchObject({
    status: "undone",
  });
  expect(reads(reviewer)).toStrictEqual(initial);
});

test("pending secondary-story suggestions retain their existing projection refusal and all proposal state", async () => {
  setSystemTime(new Date(DATE));
  const observations = observeSessions();
  const reviewer = await fixture("unlocked");
  const story = STORIES.at(-1);
  if (!story) throw new TypeError("The source declares its late story");
  const block = reviewer
    .snapshotStory(story)
    ?.blocks.find(({ text }) => text === "endnote outside");
  if (!block) throw new TypeError("The note has its real proposal anchor");
  const edit = reviewer.applyDocumentOperationsToStory({
    story,
    batch: {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [
        {
          id: "pending",
          type: "insertAfterBlock",
          blockId: block.id,
          text: "Pending note proposal.",
        },
      ],
    },
    undefinedReferences: "refuse",
  });
  expect(edit.applied).toHaveLength(1);
  expect(reviewer.exportPendingSuggestions()).toHaveLength(1);
  const original = firstSession(observations);
  const version = original.session.version;
  const before = reads(reviewer);
  const bytes = new Uint8Array(await reviewer.toBuffer());
  expectRefusal(
    () =>
      getFolioDocxComparisonAccess(reviewer).projectStories(
        "without-revision-census",
      ),
    { gap: CANONICAL_GAP.publicHeadlessSession },
  );
  expect(reads(reviewer)).toStrictEqual(before);
  expect(original.session.version).toBe(version);
  expect(new Uint8Array(await reviewer.toBuffer())).toEqual(bytes);
  expect(observations.slice(1).flatMap(({ disposals }) => disposals)).toEqual(
    [],
  );
});

for (const decision of ["accept", "reject"] as const)
  for (const scope of ["single", "all"] as const)
    test(`staged final section ${decision}/${scope} resolves through one source calculation and exact semantic inverses`, async () => {
      const observations = observeSessions();
      const reviewer = await fixture("unlocked");
      const original = firstSession(observations);
      const unsubscribe = original.session.onCommitted(() => {});
      const before = cloneDocumentWithParagraphPropertySources(
        original.session.document,
      );
      const previous =
        getFolioDocxComparisonAccess(reviewer).finalSectionProperties();
      if (!previous)
        throw new TypeError("The actual package has its final section");
      const target = {
        ...previous,
        pageWidth: 15840,
        pageHeight: 12240,
        orientation: "landscape",
      } as const;
      const access = getFolioDocxComparisonAccess(reviewer);
      const revisionId = 9000;
      expect(
        access.stageFinalSectionProperties({
          target,
          previous,
          revision: { idSeed: revisionId, date: DATE },
        }),
      ).toBe(true);
      expect(reviewer.getChanges().some(({ id }) => id === revisionId)).toBe(
        true,
      );
      const staged = cloneDocumentWithParagraphPropertySources(
        original.session.document,
      );
      const stageBatch = original.batches.at(-1);
      if (!stageBatch)
        throw new TypeError("Staging committed its real operation");
      expect(
        applyDocumentOps(staged, stageBatch.inverse).unwrap().document,
      ).toStrictEqual(before);
      expect(
        applyDocumentOps(
          staged,
          JSON.parse(JSON.stringify(stageBatch.inverse)),
        ).unwrap().document,
      ).toStrictEqual(before);
      if (scope === "single") {
        expect(
          decision === "accept"
            ? reviewer.acceptChange(revisionId)
            : reviewer.rejectChange(revisionId),
        ).toBe(true);
      } else {
        expect(
          decision === "accept" ? reviewer.acceptAll() : reviewer.rejectAll(),
        ).toEqual({ resolved: 6, skipped: [] });
      }
      const expected = decision === "accept" ? target : previous;
      expect(access.finalSectionProperties()).toStrictEqual(expected);
      const resolved = cloneDocumentWithParagraphPropertySources(
        original.session.document,
      );
      const resolution = original.batches.at(-1);
      if (!resolution || resolution === stageBatch)
        throw new TypeError("Review committed a distinct real operation");
      expect(
        applyDocumentOps(resolved, resolution.inverse).unwrap().document,
      ).toStrictEqual(staged);
      expect(
        applyDocumentOps(
          resolved,
          JSON.parse(JSON.stringify(resolution.inverse)),
        ).unwrap().document,
      ).toStrictEqual(staged);
      expect(
        applyDocumentOps(resolved, [
          ...resolution.inverse,
          ...stageBatch.inverse,
        ]).unwrap().document,
      ).toStrictEqual(before);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session: "canonical" },
      );
      expect(
        getFolioDocxComparisonAccess(reopened).finalSectionProperties(),
      ).toStrictEqual(expected);
      if (scope === "all")
        for (const story of STORIES)
          expect(
            reopened.readReviewedStory({ story, view: "current-markup" })
              ?.changes,
          ).toEqual([]);
      unsubscribe();
    });
