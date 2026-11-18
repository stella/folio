import { panic, Result } from "better-result";
import {
  afterEach,
  beforeEach,
  expect,
  setSystemTime,
  spyOn,
  test,
} from "bun:test";

import * as canonicalSession from "../canonical/canonicalSession";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

beforeEach(() => setSystemTime(new Date("2026-01-02T03:04:05Z")));
afterEach(() => setSystemTime());

const STORIES = [
  { type: "main" },
  { type: "header", relationshipId: "rIdHeader" },
  { type: "footer", relationshipId: "rIdFooter" },
  { type: "footnote", noteId: 7 },
  { type: "endnote", noteId: 8 },
] as const;
type Story = (typeof STORIES)[number];

const fixture = async () => {
  const content = () => [
    paragraph([run("Original words.")]),
    paragraph([run("Tail.")]),
  ];
  const document = fromMarkdown("Body.");
  document.package.document.content = [
    ...content(),
    paragraph([
      {
        type: "run",
        content: [
          { type: "footnoteRef", id: 7 },
          { type: "endnoteRef", id: 8 },
        ],
      },
    ]),
  ];
  document.package.document.finalSectionProperties = {
    headerReferences: [{ type: "default", rId: "rIdHeader" }],
    footerReferences: [{ type: "default", rId: "rIdFooter" }],
  };
  document.package.headers = new Map([
    [
      "rIdHeader",
      { type: "header", hdrFtrType: "default", content: content() },
    ],
  ]);
  document.package.footers = new Map([
    [
      "rIdFooter",
      { type: "footer", hdrFtrType: "default", content: content() },
    ],
  ]);
  document.package.footnotes = [
    { type: "footnote", id: 7, content: content() },
  ];
  document.package.endnotes = [{ type: "endnote", id: 8, content: content() }];
  return (await prepareCanonicalDocxInput(await createDocx(document)))
    .unwrap()
    .slice().buffer;
};

const open = (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes.slice(0), {
    session: "canonical",
    author: "Reviewer",
  });

const reads = (reviewer: FolioDocxReviewer) => ({
  stories: reviewer
    .listStories()
    .map(({ handle }) => reviewer.readStory(handle)),
  changes: STORIES.map(
    (story) =>
      reviewer.readReviewedStory({ story, view: "current-markup" })?.changes,
  ),
  comments: reviewer.getComments(),
});

const apply = (
  reviewer: FolioDocxReviewer,
  story: Story,
  operation: FolioDocumentOperation,
) =>
  reviewer.applyDocumentOperationsToStory({
    story,
    batch: {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [operation],
    },
    undefinedReferences: "refuse",
  });

const anchor = (reviewer: FolioDocxReviewer, story: Story) => {
  const block = reviewer.snapshotStory(story)?.blocks.at(0);
  if (!block) panic("The fixture story must have a first paragraph.");
  return block.id;
};

for (const story of STORIES)
  for (const kind of ["insert", "replace", "delete", "format"] as const)
    test(`${story.type} ${kind}: adoption and rejection preserve suggestion ownership`, async () => {
      const bytes = await fixture();
      const reviewer = await open(bytes);
      const before = reads(reviewer);
      const blockId = anchor(reviewer, story);
      const operations = {
        insert: {
          id: "proposal",
          type: "insertAfterBlock",
          blockId,
          text: "Proposed words.",
          comment: { text: "Proposed comment." },
        },
        replace: {
          id: "proposal",
          type: "replaceBlock",
          blockId,
          text: "Replacement words.",
        },
        delete: { id: "proposal", type: "deleteBlock", blockId },
        format: {
          id: "proposal",
          type: "setBlockParagraphProperties",
          blockId,
          properties: { alignment: "right" },
        },
      } as const satisfies Record<typeof kind, FolioDocumentOperation>;
      const receipt = apply(reviewer, story, operations[kind]);
      expect(receipt.applied).toHaveLength(1);
      expect(receipt.skipped).toEqual([]);
      const staged = reads(reviewer);
      const records = reviewer.exportPendingSuggestions();
      expect(records).toHaveLength(1);
      const proposal = records.at(0);
      if (!proposal) panic("An applied suggestion must have an export record.");

      // Bulk acceptance leaves proposals staged in every story.
      expect(reviewer.acceptAll()).toEqual({ resolved: 0, skipped: [] });
      expect(reads(reviewer)).toEqual(staged);
      expect(reviewer.exportPendingSuggestions()).toEqual(records);
      const saved = await reviewer.toBuffer();
      const restaged = await open(saved);
      expect(reads(restaged)).toEqual(before);
      expect(
        restaged.loadPendingSuggestions(records).map(({ status }) => status),
      ).toEqual(["restaged"]);
      expect(reads(restaged)).toEqual(staged);
      expect(restaged.rejectSuggestion(proposal.suggestionId)).toBe(true);
      expect(reads(restaged)).toEqual(before);
      expect(restaged.exportPendingSuggestions()).toEqual([]);

      expect(
        reviewer.acceptSuggestion(proposal.suggestionId, { author: "Adopter" }),
      ).toBe(true);
      expect(reviewer.exportPendingSuggestions()).toEqual([]);
      expect(reviewer.acceptAll().skipped).toEqual([]);
      expect(
        reviewer.readReviewedStory({ story, view: "current-markup" })?.changes,
      ).toEqual([]);
      const direct = await open(bytes);
      const result = direct.applyDocumentOperationsToStory({
        story,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: [operations[kind]],
        },
        undefinedReferences: "refuse",
      });
      expect(result.applied).toHaveLength(1);
      expect(reviewer.readStory(story)?.text).toEqual(
        direct.readStory(story)?.text,
      );
      expect(
        reviewer
          .getComments()
          .map(({ text, anchoredText }) => ({ text, anchoredText })),
      ).toEqual(
        direct
          .getComments()
          .map(({ text, anchoredText }) => ({ text, anchoredText })),
      );
      expect(reads(await open(await reviewer.toBuffer()))).toEqual(
        reads(reviewer),
      );
    });

const stageEveryStory = async () => {
  const reviewer = await open(await fixture());
  for (const story of STORIES) {
    const result = reviewer.applyDocumentOperationsToStory({
      story,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          {
            id: `ordinary-${story.type}`,
            type: "commentOnBlock",
            blockId: anchor(reviewer, story),
            comment: { text: `Keep ${story.type}.` },
          },
        ],
      },
      undefinedReferences: "refuse",
    });
    expect(result.applied).toHaveLength(1);
  }
  const before = reads(reviewer);
  for (const story of STORIES) {
    const result = apply(reviewer, story, {
      id: `proposal-${story.type}`,
      type: "insertAfterBlock",
      blockId: anchor(reviewer, story),
      text: `Added ${story.type}.`,
      comment: { text: `Comment ${story.type}.` },
    });
    expect(result.applied).toHaveLength(1);
  }
  return { reviewer, before };
};

const observeResolution = (mode: "observe" | "refuse") => {
  const create = canonicalSession.createCanonicalSession;
  let calls = 0;
  const factory = spyOn(
    canonicalSession,
    "createCanonicalSession",
  ).mockImplementation((...args) => {
    const created = create(...args);
    if (created.isErr()) return created;
    const session = created.value;
    const prepare = session.prepareResolve.bind(session);
    session.prepareResolve = (...input) => {
      calls++;
      if (mode === "refuse")
        return Result.err(
          new canonicalSession.CanonicalSessionError({
            reason: "refused",
            ground: { gap: CANONICAL_GAP.publicHeadlessSession },
            message: "Refused staged resolution.",
          }),
        );
      return prepare(...input);
    };
    return created;
  });
  return { calls: () => calls, restore: () => factory.mockRestore() };
};

test("bulk rejection resolves every story and proposed comment in one prepared commit", async () => {
  const observer = observeResolution("observe");
  try {
    const { reviewer, before } = await stageEveryStory();
    expect(observer.calls()).toBe(0);
    const result = reviewer.rejectAll();
    expect(result.resolved).toBeGreaterThan(0);
    expect(result.skipped).toEqual([]);
    expect(observer.calls()).toBe(1);
    expect(reads(reviewer)).toEqual(before);
    expect(reviewer.exportPendingSuggestions()).toEqual([]);
  } finally {
    observer.restore();
  }
});

test("a refused bulk preparation leaves every staged story and export record intact", async () => {
  const observer = observeResolution("refuse");
  try {
    const { reviewer } = await stageEveryStory();
    const before = reads(reviewer);
    const records = reviewer.exportPendingSuggestions();
    expect(observer.calls()).toBe(0);
    expect(() => reviewer.rejectAll()).toThrow("Refused staged resolution.");
    expect(observer.calls()).toBe(1);
    expect(reads(reviewer)).toEqual(before);
    expect(reviewer.exportPendingSuggestions()).toEqual(records);
  } finally {
    observer.restore();
  }
});
