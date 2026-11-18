import { panic } from "better-result";
import { expect, spyOn, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOp,
  applyDocumentOps,
  findStoryBody,
  normalizeForOps,
  OP_STORIES,
  type OpStory,
  pendingSuggestionIdsIn,
} from "@stll/docx-core/ops";

import * as canonicalSession from "../canonical/canonicalSession";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { schema } from "../prosemirror/schema";
import { CANONICAL_RULE } from "../types/canonicalCapabilities";
import type { Paragraph } from "../types/document";
import { readModel } from "../types/documentModel";
import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "./headless";

const PROPOSAL = "mixed-lock-proposal";
const STORIES = [
  {
    story: OP_STORIES.MAIN,
    handle: { type: "main" },
    held: "10000001",
    free: "10000002",
  },
  {
    story: { kind: "header", rId: "rIdHeader" },
    handle: { type: "header", relationshipId: "rIdHeader" },
    held: "20000001",
    free: "20000002",
  },
  {
    story: { kind: "footer", rId: "rIdFooter" },
    handle: { type: "footer", relationshipId: "rIdFooter" },
    held: "30000001",
    free: "30000002",
  },
  {
    story: { kind: "footnote", id: 7 },
    handle: { type: "footnote", noteId: 7 },
    held: "40000001",
    free: "40000002",
  },
  {
    story: { kind: "endnote", id: 8 },
    handle: { type: "endnote", noteId: 8 },
    held: "50000001",
    free: "50000002",
  },
] as const satisfies readonly {
  story: OpStory;
  handle: Parameters<FolioDocxReviewer["readStory"]>[0];
  held: string;
  free: string;
}[];

const content = (
  { held, free }: (typeof STORIES)[number],
  lock?: "sdtLocked",
): Paragraph[] => [
  {
    type: "paragraph",
    paraId: held,
    content: [
      {
        type: "inlineSdt",
        properties: {
          id: 9,
          sdtType: "richText",
          ...(lock === undefined ? {} : { lock }),
        },
        content: [{ type: "run", content: [{ type: "text", text: "Held" }] }],
      },
    ],
  },
  {
    type: "paragraph",
    paraId: free,
    content: [{ type: "run", content: [{ type: "text", text: "Free" }] }],
  },
];

const fixture = async (lock?: "sdtLocked") => {
  const document = fromMarkdown("Body.");
  document.package.document.content = [
    ...content(STORIES[0], lock),
    {
      type: "paragraph",
      paraId: "60000001",
      content: [
        {
          type: "run",
          content: [
            { type: "footnoteRef", id: 7 },
            { type: "endnoteRef", id: 8 },
          ],
        },
      ],
    },
  ];
  document.package.document.finalSectionProperties = {
    headerReferences: [{ type: "default", rId: "rIdHeader" }],
    footerReferences: [{ type: "default", rId: "rIdFooter" }],
  };
  document.package.headers = new Map([
    [
      "rIdHeader",
      {
        type: "header",
        hdrFtrType: "default",
        content: content(STORIES[1], lock),
      },
    ],
  ]);
  document.package.footers = new Map([
    [
      "rIdFooter",
      {
        type: "footer",
        hdrFtrType: "default",
        content: content(STORIES[2], lock),
      },
    ],
  ]);
  document.package.footnotes = [
    { type: "footnote", id: 7, content: content(STORIES[3], lock) },
  ];
  document.package.endnotes = [
    { type: "endnote", id: 8, content: content(STORIES[4], lock) },
  ];
  return (await prepareCanonicalDocxInput(await createDocx(document)))
    .unwrap()
    .slice().buffer;
};

/** Produce the proposal first; then another package edit locks one of its two changes. */
const proposed = async (entry: (typeof STORIES)[number]) => {
  const bytes = await fixture();
  let document = normalizeForOps(
    await parseDocx(bytes, { preloadFonts: false }),
  );
  for (const [index, blockId] of [entry.held, entry.free].entries())
    document = applyDocumentOp(document, {
      type: "insertText",
      at: { story: entry.story, blockId, offset: 1 },
      text: " proposed",
      runProps: {},
      revision: {
        id: index + 1,
        author: "Proposer",
        date: "2026-01-02T03:04:05Z",
        suggestionId: PROPOSAL,
      },
    }).unwrap().document;
  document = applyDocumentOp(document, {
    type: "createComment",
    comment: {
      id: 7,
      author: "Proposer",
      suggestionId: PROPOSAL,
      content: [
        {
          type: "paragraph",
          paraId: "70000001",
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Keep while held." }],
            },
          ],
        },
      ],
    },
    anchor: {
      kind: "point",
      at: { story: entry.story, blockId: entry.free, offset: 0 },
    },
  }).unwrap().document;
  const paragraph = findStoryBody(document, entry.story)?.content.find(
    (block) => block.type === "paragraph" && block.paraId === entry.held,
  );
  if (paragraph?.type !== "paragraph")
    panic("The parsed held paragraph must exist.");
  const control = paragraph.content.find((item) => item.type === "inlineSdt");
  if (control?.type !== "inlineSdt")
    panic("The parsed held control must exist.");
  const locked = applyDocumentOp(document, {
    type: "setContentControl",
    story: entry.story,
    controlId:
      control.properties.id ?? panic("A parsed control needs its actual id."),
    blockIds: [entry.held],
    expected: control.properties,
    properties: { ...control.properties, lock: "contentLocked" },
  }).unwrap().document;
  return { bytes, unlocked: document, locked };
};

for (const entry of STORIES) {
  test(`${entry.handle.type}: one held part refuses rejection of the whole proposal atomically`, async () => {
    const { unlocked, locked } = await proposed(entry);
    for (const [document, held] of [
      [unlocked, false],
      [locked, true],
    ] as const) {
      const session = canonicalSession
        .createCanonicalSession(document)
        .unwrap();
      const state = EditorState.create({ schema, doc: session.projection.doc });
      const documentBefore = session.document;
      const original = structuredClone(readModel(documentBefore));
      expect(pendingSuggestionIdsIn(original)).toEqual(new Set([PROPOSAL]));
      const result = session.prepareSuggestion(state, {
        suggestionId: PROPOSAL,
        resolution: "reject",
        author: "Reviewer",
        date: "2026-01-02T03:04:06Z",
      });
      if (held) {
        expect(result.isErr()).toBe(true);
        if (result.isOk())
          panic("A held proposal must refuse before publication.");
        expect(result.error.ground).toEqual({
          rule: CANONICAL_RULE.contentControlLock,
        });
      } else {
        expect(result.isOk()).toBe(true);
        if (result.isErr()) panic(result.error.message);
        const shown = canonicalSession
          .publishCanonicalProjection({ state, session, commit: result.value })
          .unwrap();
        expect(shown.state.doc).toBe(session.projection.doc);
        expect(pendingSuggestionIdsIn(readModel(session.document))).toEqual(
          new Set(),
        );
        expect(session.getCommittedComments()).toEqual([]);
        continue;
      }
      expect(session.document).toBe(documentBefore);
      expect(structuredClone(readModel(session.document))).toStrictEqual(
        original,
      );
      expect(pendingSuggestionIdsIn(readModel(session.document))).toEqual(
        new Set([PROPOSAL]),
      );
      expect(session.version).toBe(0);
      expect(session.canUndo).toBe(false);
    }
  });

  test(`${entry.handle.type}: bulk rejection retains every fragment and comment of a held proposal`, async () => {
    const { bytes, locked } = await proposed(entry);
    const create = canonicalSession.createCanonicalSession;
    let owner: canonicalSession.CanonicalSession | undefined;
    // Initialize the real reviewer with the actual produced model; no resolution is stubbed.
    const factory = spyOn(
      canonicalSession,
      "createCanonicalSession",
    ).mockImplementation((_document, styles, options) => {
      const result = create(locked, styles, options);
      if (result.isOk()) owner = result.value;
      return result;
    });
    try {
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
        session: "canonical",
      });
      if (owner === undefined) panic("The reviewer must own a real session.");
      expect(
        reviewer.readReviewedStory({
          story: entry.handle,
          view: "current-markup",
        })?.changes,
      ).toHaveLength(2);
      expect(reviewer.getComments().map(({ text }) => text)).toEqual([
        "Keep while held.",
      ]);
      const sourceBefore = owner.document;
      const changesBefore = reviewer.getChanges();
      const commentsBefore = reviewer.getComments();
      expect(reviewer.rejectSuggestion(PROPOSAL)).toBe(false);
      expect(owner.document).toBe(sourceBefore);
      expect(reviewer.getChanges()).toEqual(changesBefore);
      expect(reviewer.getComments()).toEqual(commentsBefore);
      expect(pendingSuggestionIdsIn(readModel(owner.document))).toEqual(
        new Set([PROPOSAL]),
      );
      const result = reviewer.rejectAll();
      expect(result).toEqual({
        resolved: 0,
        skipped: [1, 2].map((id) => ({
          id,
          reason: "contentControlLock",
          message: expect.any(String),
        })),
      });
      expect(
        reviewer
          .readReviewedStory({ story: entry.handle, view: "current-markup" })
          ?.changes.map(({ id }) => id),
      ).toEqual([1, 2]);
      expect(reviewer.getComments().map(({ text }) => text)).toEqual([
        "Keep while held.",
      ]);
      expect(pendingSuggestionIdsIn(readModel(owner.document))).toEqual(
        new Set([PROPOSAL]),
      );
      expect(
        readModel(owner.document).package.document.comments?.map(
          ({ id, suggestionId }) => ({ id, suggestionId }),
        ),
      ).toEqual([{ id: 7, suggestionId: PROPOSAL }]);
    } finally {
      factory.mockRestore();
    }
  });
}

const publicChanges = (reviewer: FolioDocxReviewer) =>
  STORIES.flatMap(
    ({ handle }) =>
      reviewer.readReviewedStory({ story: handle, view: "current-markup" })
        ?.changes ?? [],
  );
const publicStoryText = (reviewer: FolioDocxReviewer) =>
  STORIES.map(({ handle }) => reviewer.readStory(handle)?.text);

for (const [index, entry] of STORIES.entries())
  for (const placement of ["same", "other"] as const)
    test(`${entry.handle.type} ${placement} story: a held public proposal survives bulk rejection and save/restage across stories`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await fixture("sdtLocked"),
        { session: "canonical" },
      );
      const free =
        STORIES.at(
          placement === "same" ? index : (index + 1) % STORIES.length,
        ) ?? panic("A free story must exist.");
      const ordinary =
        STORIES.at((index + 2) % STORIES.length) ??
        panic("An ordinary story must exist.");
      const apply = (
        story: (typeof STORIES)[number],
        mode: "suggested" | "tracked-changes",
        operation: Parameters<
          typeof reviewer.applyDocumentOperationsToStory
        >[0]["batch"]["operations"][number],
      ) => {
        const result = reviewer.applyDocumentOperationsToStory({
          story: story.handle,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: [operation],
          },
          undefinedReferences: "refuse",
        });
        expect(result.skipped).toEqual([]);
        expect(result.applied).toHaveLength(1);
      };
      apply(entry, "suggested", {
        id: "held-split",
        suggestionId: PROPOSAL,
        type: "splitBlock",
        blockId: entry.held,
        offset: 2,
      });
      apply(free, "suggested", {
        id: "free-proposal",
        suggestionId: PROPOSAL,
        type: "insertAfterBlock",
        blockId: free.free,
        text: "Free proposal.",
        comment: { text: "Keep this proposal comment." },
      });
      const records = reviewer.exportPendingSuggestions();
      expect(records).toHaveLength(2);
      const proposedIds = publicChanges(reviewer)
        .map(({ id }) => id)
        .toSorted((left, right) => left - right);
      expect(proposedIds.length).toBeGreaterThan(1);
      const before = publicStoryText(reviewer);
      const comments = reviewer.getComments();
      apply(ordinary, "tracked-changes", {
        id: "ordinary-free",
        type: "insertAfterBlock",
        blockId: ordinary.free,
        text: "Ordinary free change.",
      });
      const ordinaryIds = publicChanges(reviewer)
        .filter(({ id }) => !proposedIds.includes(id))
        .map(({ id }) => id);
      expect(ordinaryIds.length).toBeGreaterThan(0);
      const result = reviewer.rejectAll();
      expect(result.resolved).toBe(ordinaryIds.length);
      expect(
        result.skipped
          .map(({ id }) => id)
          .toSorted((left, right) => left - right),
      ).toEqual(proposedIds);
      expect(
        result.skipped.every(({ reason }) => reason === "contentControlLock"),
      ).toBe(true);
      expect(publicStoryText(reviewer)).toEqual(before);
      expect(reviewer.getComments()).toEqual(comments);
      expect(reviewer.exportPendingSuggestions()).toEqual(records);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session: "canonical" },
      );
      expect(
        reopened.loadPendingSuggestions(records).map(({ status }) => status),
      ).toEqual(["restaged", "restaged"]);
      expect(publicStoryText(reopened)).toEqual(before);
      expect(reopened.exportPendingSuggestions()).toEqual(records);
      expect(
        reopened
          .getComments()
          .map(({ text, anchoredText }) => ({ text, anchoredText })),
      ).toEqual(
        comments.map(({ text, anchoredText }) => ({ text, anchoredText })),
      );
    });

for (const ordinary of ["present", "absent"] as const)
  test(`bulk rejection resolves a free final section beside a held cross-story proposal, ordinary change ${ordinary}`, async () => {
    const create = canonicalSession.createCanonicalSession;
    let owner: canonicalSession.CanonicalSession | undefined;
    const factory = spyOn(
      canonicalSession,
      "createCanonicalSession",
    ).mockImplementation((...args) => {
      const result = create(...args);
      if (result.isOk()) owner ??= result.value;
      return result;
    });
    try {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await fixture("sdtLocked"),
        { session: "canonical" },
      );
      if (!owner) panic("A canonical reviewer must own its source session.");
      const apply = (
        story: (typeof STORIES)[number],
        mode: "suggested" | "tracked-changes",
        operation: Parameters<
          typeof reviewer.applyDocumentOperationsToStory
        >[0]["batch"]["operations"][number],
      ) => {
        const result = reviewer.applyDocumentOperationsToStory({
          story: story.handle,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: [operation],
          },
          undefinedReferences: "refuse",
        });
        expect(result.skipped).toEqual([]);
        expect(result.applied).toHaveLength(1);
      };
      apply(STORIES[0], "suggested", {
        id: "held-split",
        suggestionId: PROPOSAL,
        type: "splitBlock",
        blockId: STORIES[0].held,
        offset: 2,
      });
      apply(STORIES[1], "suggested", {
        id: "free-proposal",
        suggestionId: PROPOSAL,
        type: "insertAfterBlock",
        blockId: STORIES[1].free,
        text: "Free proposal.",
        comment: { text: "Keep this proposal comment." },
      });
      const records = reviewer.exportPendingSuggestions();
      expect(records).toHaveLength(2);
      const proposedIds = publicChanges(reviewer)
        .map(({ id }) => id)
        .toSorted((left, right) => left - right);
      expect(proposedIds.length).toBeGreaterThan(1);
      const proposalText = publicStoryText(reviewer);
      const comments = reviewer.getComments();
      if (ordinary === "present")
        apply(STORIES[4], "tracked-changes", {
          id: "ordinary-free",
          type: "insertAfterBlock",
          blockId: STORIES[4].free,
          text: "Ordinary free change.",
        });
      const ordinaryIds = publicChanges(reviewer)
        .filter(({ id }) => !proposedIds.includes(id))
        .map(({ id }) => id);
      expect(ordinaryIds.length > 0).toBe(ordinary === "present");
      const access = getFolioDocxComparisonAccess(reviewer);
      const previous =
        access.finalSectionProperties() ??
        panic("The parsed package has its final section.");
      const target = {
        ...previous,
        pageWidth: 15840,
        pageHeight: 12240,
        orientation: "landscape",
      } as const;
      expect(
        access.stageFinalSectionProperties({
          previous,
          target,
          revision: { idSeed: 9000, date: "2026-01-02T03:04:05Z" },
        }),
      ).toBe(true);
      const staged = cloneDocumentWithParagraphPropertySources(owner.document);
      const committed: canonicalSession.CanonicalCommittedBatch[] = [];
      const unsubscribe = owner.onCommitted((batch) => {
        committed.push(batch);
      });
      const prepare = spyOn(owner, "prepareResolve");
      const result = reviewer.rejectAll();
      expect(prepare).toHaveBeenCalledTimes(1);
      expect(committed).toHaveLength(1);
      unsubscribe();
      expect(result.resolved).toBe(ordinaryIds.length + 1);
      expect(
        result.skipped
          .map(({ id }) => id)
          .toSorted((left, right) => left - right),
      ).toEqual(proposedIds);
      expect(
        result.skipped.every(({ reason }) => reason === "contentControlLock"),
      ).toBe(true);
      expect(access.finalSectionProperties()).toStrictEqual(previous);
      expect(publicStoryText(reviewer)).toEqual(proposalText);
      expect(reviewer.getComments()).toEqual(comments);
      expect(reviewer.exportPendingSuggestions()).toEqual(records);
      expect(
        publicChanges(reviewer)
          .map(({ id }) => id)
          .toSorted((left, right) => left - right),
      ).toEqual(proposedIds);
      const resolution =
        committed.at(0) ?? panic("Bulk rejection commits once.");
      const resolved = cloneDocumentWithParagraphPropertySources(
        owner.document,
      );
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
        applyDocumentOps(
          staged,
          JSON.parse(JSON.stringify(resolution.ops)),
        ).unwrap().document,
      ).toStrictEqual(resolved);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session: "canonical" },
      );
      expect(
        getFolioDocxComparisonAccess(reopened).finalSectionProperties(),
      ).toStrictEqual(previous);
      expect(
        reopened.loadPendingSuggestions(records).map(({ status }) => status),
      ).toEqual(["restaged", "restaged"]);
      expect(publicStoryText(reopened)).toEqual(proposalText);
      expect(reopened.exportPendingSuggestions()).toEqual(records);
      expect(
        reopened
          .getComments()
          .map(({ text, anchoredText }) => ({ text, anchoredText })),
      ).toEqual(
        comments.map(({ text, anchoredText }) => ({ text, anchoredText })),
      );
    } finally {
      factory.mockRestore();
    }
  });
