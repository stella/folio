import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import { expectParagraphBlock } from "../../../../test/paragraphBlock";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationBatch,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { type FolioDocumentStoryHandle, FolioDocxReviewer } from "./headless";

const STORIES = {
  main: { type: "main" },
  header: { type: "header", relationshipId: "rIdHeader1" },
  footer: { type: "footer", relationshipId: "rIdFooter1" },
  footnote: { type: "footnote", noteId: 1 },
  endnote: { type: "endnote", noteId: 1 },
} as const satisfies Record<
  FolioDocumentStoryHandle["type"],
  FolioDocumentStoryHandle
>;

const SESSIONS: readonly HeadlessReviewerSession[] = ["canonical"];

const SESSION_STORIES = SESSIONS.flatMap((session) =>
  Object.values(STORIES).map(({ type }) => ({ session, type })),
);

const open = (
  session: HeadlessReviewerSession,
  bytes: Parameters<typeof FolioDocxReviewer.fromBuffer>[0],
) => FolioDocxReviewer.fromBuffer(bytes, { session });

const PARTS = {
  main: "word/document.xml",
  header: "word/header1.xml",
  footer: "word/footer1.xml",
  footnote: "word/footnotes.xml",
  endnote: "word/endnotes.xml",
} as const satisfies Record<keyof typeof STORIES, string>;

const fixture = async (
  length: number,
  referencePosition: "leading" | "trailing" = "trailing",
) => {
  const document = createEmptyDocument();
  const paragraphs = (prefix: number) =>
    Array.from(
      { length },
      (_, index) =>
        ({
          type: "paragraph",
          paraId: (prefix + index).toString(16),
          formatting: { alignment: "left" },
          content: [
            {
              type: "run",
              content: [{ type: "text", text: `Clause ${index}.` }],
            },
          ],
        }) satisfies Paragraph,
    );
  document.package.document.content = paragraphs(0x10000000);
  const references = {
    type: "paragraph",
    paraId: "20000000",
    content: [
      {
        type: "run",
        content: [
          { type: "footnoteRef", id: 1 },
          { type: "endnoteRef", id: 1 },
        ],
      },
    ],
  } satisfies Paragraph;
  if (referencePosition === "leading")
    document.package.document.content.unshift(references);
  else document.package.document.content.push(references);
  document.package.headers = new Map([
    [
      STORIES.header.relationshipId,
      {
        type: "header",
        hdrFtrType: "default",
        content: paragraphs(0x30000000),
      },
    ],
  ]);
  document.package.footers = new Map([
    [
      STORIES.footer.relationshipId,
      {
        type: "footer",
        hdrFtrType: "default",
        content: paragraphs(0x40000000),
      },
    ],
  ]);
  document.package.footnotes = [
    { type: "footnote", id: 1, content: paragraphs(0x50000000) },
  ];
  document.package.endnotes = [
    { type: "endnote", id: 1, content: paragraphs(0x60000000) },
  ];
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: STORIES.header.relationshipId }],
    footerReferences: [{ type: "default", rId: STORIES.footer.relationshipId }],
  };
  return await createDocx(document);
};

type BlockOptions = {
  reviewer: FolioDocxReviewer;
  story: FolioDocumentStoryHandle;
  index: number;
};

const block = ({ reviewer, story, index }: BlockOptions) => {
  const target = reviewer.snapshotStory(story)?.blocks.at(index);
  if (!target)
    panic("The story capacity fixture lost its paragraph", { story, index });
  return target;
};

type ApplyOptions = {
  reviewer: FolioDocxReviewer;
  story: FolioDocumentStoryHandle;
  operation: FolioDocumentOperationBatch["operations"][number];
};

const apply = ({ reviewer, story, operation }: ApplyOptions) =>
  reviewer.applyDocumentOperationsToStory({
    undefinedReferences: "refuse",
    story,
    batch: {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      atomic: true,
      operations: [operation],
    },
  });

type CapacityOptions = {
  reviewer: FolioDocxReviewer;
  kind: keyof typeof STORIES;
  occupiedParaId?: string;
};

const capacity = async ({
  reviewer,
  kind,
  occupiedParaId,
}: CapacityOptions) => {
  const zip = await ZipArchive.read(await reviewer.toBuffer());
  const part = zip.entry(PARTS[kind]);
  if (!part)
    panic("The story capacity fixture lost its package part", { kind });
  const xml = await part.text();
  const paragraphs = xml.match(/<w:p\b[\s\S]*?<\/w:p>/gu) ?? [];
  expect(paragraphs.length).toBeGreaterThan(0);
  for (const paragraph of paragraphs) {
    expect(
      paragraph.match(/<w:pPrChange\b/gu)?.length ?? 0,
    ).toBeLessThanOrEqual(1);
  }
  if (occupiedParaId !== undefined) {
    const target = paragraphs.find((paragraph) =>
      paragraph.includes(`w14:paraId="${occupiedParaId}"`),
    );
    expect(target).toBeDefined();
    expect(target?.match(/<w:pPrChange\b/gu)).toHaveLength(1);
  }
};

type CapacityCase = {
  session: HeadlessReviewerSession;
  kind: keyof typeof STORIES;
  attempts: number;
  length: number;
  pendingAt: "first" | "survivor";
  merge: "explicit" | "carried";
};

const checkCapacity = async ({
  session,
  kind,
  attempts,
  length,
  pendingAt,
  merge,
}: CapacityCase) => {
  const story = STORIES[kind];
  const reviewer = await open(session, await fixture(length));
  // Deleted intermediate marks move the formatting destination beyond the next paragraph.
  for (let index = length - 2; index >= 1; index--) {
    expect(
      apply({
        reviewer,
        story,
        operation: {
          id: `chain-${index}`,
          type: "mergeBlockWithNext",
          blockId: block({ reviewer, story, index }).id,
        },
      }).status,
    ).toBe("committed");
  }
  const targetIndex = pendingAt === "first" ? 0 : length - 1;
  for (let index = 0; index < attempts; index++) {
    const result = apply({
      reviewer,
      story,
      operation: {
        id: `property-${index}`,
        type: "setBlockParagraphProperties",
        blockId: block({ reviewer, story, index: targetIndex }).id,
        properties: { alignment: index % 2 === 0 ? "center" : "right" },
      },
    });
    // A later change by the same author folds into the pending one, which
    // keeps the original properties.
    expect(result.status).toBe("committed");
    await capacity({ reviewer, kind });
  }
  if (attempts > 0) {
    // The two-paragraph regression must actually occupy the writer's slot.
    await capacity({
      reviewer,
      kind,
      occupiedParaId: block({ reviewer, story, index: targetIndex }).id,
    });
  }
  const result = apply({
    reviewer,
    story,
    operation: {
      id: "formatting-merge",
      type: "mergeBlockWithNext",
      blockId: block({ reviewer, story, index: 0 }).id,
      separator: " ",
      ...(merge === "explicit"
        ? { mergedParagraphProperties: { alignment: "both" as const } }
        : {}),
    },
  });
  expect(result.status).toBe("committed");
  await capacity({ reviewer, kind });
  const saved = await reviewer.toBuffer();
  const rejecting = await open(session, saved);
  rejecting.rejectAll();
  const reopenedRejected = await open(session, await rejecting.toBuffer());
  expect(
    reopenedRejected
      .snapshotStory(story)
      ?.blocks.slice(0, length)
      .map((snapshotBlock) => {
        const paragraph = expectParagraphBlock(snapshotBlock);
        return {
          text: paragraph.text,
          directAlignment: paragraph.directAlignment,
        };
      }),
  ).toEqual(
    Array.from({ length }, (_, index) => ({
      text: `Clause ${index}.`,
      directAlignment: "left",
    })),
  );
  const accepting = await open(session, saved);
  accepting.acceptAll();
  const reopenedAccepted = await open(session, await accepting.toBuffer());
  expect(
    reopenedAccepted.readReviewedStory({ story, view: "current-markup" })
      ?.changes,
  ).toEqual([]);
  const aligned = reopenedAccepted
    .snapshotStory(story)
    ?.blocks.find(({ text }) => text.includes(`Clause ${targetIndex}.`));
  // The last property change that applied.
  const pendingAlignment = (attempts - 1) % 2 === 0 ? "center" : "right";
  let mergedAlignment = "left";
  if (attempts > 0 && pendingAt === "first") mergedAlignment = pendingAlignment;
  if (merge === "explicit") mergedAlignment = "both";
  expect(expectParagraphBlock(aligned).directAlignment).toBe(mergedAlignment);
};

// The old merge guard inspected only the first paragraph; its formatting writer
// targeted the survivor. Vary both ownership and the number of attempted revisions.
test.each(SESSION_STORIES)(
  "$session: formatting merge refuses the pending revision in a $type survivor",
  async ({ session, type }) => {
    await checkCapacity({
      session,
      kind: type,
      attempts: 1,
      length: 2,
      pendingAt: "survivor",
      merge: "explicit",
    });
  },
);

test.each(SESSION_STORIES)(
  "$session: plain merge carries formatting using the pending $type survivor receipt",
  async ({ session, type }) => {
    await checkCapacity({
      session,
      kind: type,
      attempts: 1,
      length: 2,
      pendingAt: "survivor",
      merge: "carried",
    });
  },
);

test.each(SESSION_STORIES)(
  "$session: $type formatting sequences preserve paragraph revision capacity",
  async ({ session, type }) => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 0, max: 8 }),
        fc.integer({ min: 2, max: 5 }),
        fc.constantFrom("first" as const, "survivor" as const),
        fc.constantFrom("explicit" as const, "carried" as const),
        async (attempts, length, pendingAt, merge) => {
          await checkCapacity({
            session,
            kind: type,
            attempts,
            length,
            pendingAt,
            merge,
          });
        },
      ),
      {
        numRuns: 20,
        id: "$type formatting sequences preserve paragraph revision capacity",
      },
    );
  },
  propertyTestTimeout(120_000),
);

test.each(SESSION_STORIES)(
  "$session: equal-property plain merge preserves an occupied $type survivor revision",
  async ({ session, type }) => {
    const story = STORIES[type];
    const reviewer = await open(session, await fixture(2));
    const first = block({ reviewer, story, index: 0 });
    const survivor = block({ reviewer, story, index: 1 });
    expect(
      reviewer.applyDocumentOperationsToStory({
        undefinedReferences: "refuse",
        story,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: [
            {
              id: "source-format",
              type: "setBlockParagraphProperties",
              blockId: first.id,
              properties: { alignment: "center" },
            },
          ],
        },
      }).status,
    ).toBe("committed");
    expect(
      apply({
        reviewer,
        story,
        operation: {
          id: "survivor-format",
          type: "setBlockParagraphProperties",
          blockId: survivor.id,
          properties: { alignment: "center" },
        },
      }).status,
    ).toBe("committed");
    expect(
      apply({
        reviewer,
        story,
        operation: {
          id: "equal-merge",
          type: "mergeBlockWithNext",
          blockId: first.id,
          separator: " ",
        },
      }).status,
    ).toBe("committed");
    await capacity({ reviewer, kind: type });
    const saved = await reviewer.toBuffer();
    const accepting = await open(session, saved);
    accepting.acceptAll();
    const accepted = await open(session, await accepting.toBuffer());
    expect(accepted.snapshotStory(story)?.blocks.at(0)).toMatchObject({
      text: "Clause 0. Clause 1.",
      directAlignment: "center",
    });
    const rejecting = await open(session, saved);
    rejecting.rejectAll();
    const rejected = await open(session, await rejecting.toBuffer());
    expect(
      rejected
        .snapshotStory(story)
        ?.blocks.slice(0, 2)
        .map(({ text, directAlignment }) => ({ text, directAlignment })),
    ).toEqual([
      { text: "Clause 0.", directAlignment: "center" },
      { text: "Clause 1.", directAlignment: "left" },
    ]);
  },
);

test.each(SESSION_STORIES)(
  "$session: plain merge retracts an inserted $type break with a pending property revision",
  async ({ session, type }) => {
    const story = STORIES[type];
    const reviewer = await open(session, await fixture(2));
    expect(
      apply({
        reviewer,
        story,
        operation: {
          id: "split",
          type: "splitBlock",
          blockId: block({ reviewer, story, index: 0 }).id,
          offset: 7,
        },
      }).status,
    ).toBe("committed");
    const first = block({ reviewer, story, index: 0 });
    expect(
      apply({
        reviewer,
        story,
        operation: {
          id: "property",
          type: "setBlockParagraphProperties",
          blockId: first.id,
          properties: { alignment: "center" },
        },
      }).status,
    ).toBe("committed");
    await capacity({ reviewer, kind: type, occupiedParaId: first.id });
    expect(
      apply({
        reviewer,
        story,
        operation: {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: first.id,
          separator: " ",
        },
      }).status,
    ).toBe("committed");
    await capacity({ reviewer, kind: type });
    const saved = await reviewer.toBuffer();
    const accepting = await open(session, saved);
    accepting.acceptAll();
    expect(
      expectParagraphBlock(accepting.snapshotStory(story)?.blocks.at(0))
        .directAlignment,
    ).toBe("center");
    const rejecting = await open(session, saved);
    rejecting.rejectAll();
    expect(
      rejecting
        .snapshotStory(story)
        ?.blocks.slice(0, 2)
        .map((snapshotBlock) => {
          const paragraph = expectParagraphBlock(snapshotBlock);
          return {
            text: paragraph.text,
            directAlignment: paragraph.directAlignment,
          };
        }),
    ).toEqual([
      { text: "Clause 0.", directAlignment: "left" },
      { text: "Clause 1.", directAlignment: "left" },
    ]);
  },
);

test.each(SESSION_STORIES)(
  "$session: explicit merge reserves a deferred inserted-break carry in a $type story",
  async ({ session, type }) => {
    const story = STORIES[type];
    const bytes = await fixture(2, "leading");
    const prepare = async () => {
      const reviewer = await open(session, bytes);
      const sourceIndex = type === "main" ? 1 : 0;
      const source = block({ reviewer, story, index: sourceIndex });
      const survivor = block({ reviewer, story, index: sourceIndex + 1 });
      expect(
        apply({
          reviewer,
          story,
          operation: {
            id: "insert-carrier",
            type: "insertBeforeBlock",
            blockId: survivor.id,
            text: "Inserted carry.",
          },
        }).status,
      ).toBe("committed");
      const carrier = block({ reviewer, story, index: sourceIndex + 1 });
      expect(
        reviewer.applyDocumentOperationsToStory({
          undefinedReferences: "refuse",
          story,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "direct",
            operations: [
              {
                id: "carrier-format",
                type: "setBlockParagraphProperties",
                blockId: carrier.id,
                properties: { alignment: "center" },
              },
            ],
          },
        }).status,
      ).toBe("committed");
      return { reviewer, source, survivor };
    };
    const batch = await prepare();
    const sequential = await prepare();
    const operations = ({
      source,
      survivor,
    }: Awaited<ReturnType<typeof prepare>>) =>
      [
        {
          id: "delete-final",
          type: "deleteBlock",
          blockId: survivor.id,
        },
        {
          id: "explicit-merge",
          type: "mergeBlockWithNext",
          blockId: source.id,
          mergedParagraphProperties: { alignment: "right" },
        },
      ] as const satisfies FolioDocumentOperationBatch["operations"];
    const sequentialOperations = operations(sequential);
    expect(
      apply({
        reviewer: sequential.reviewer,
        story,
        operation: sequentialOperations[0],
      }).status,
    ).toBe("committed");
    await capacity({
      reviewer: sequential.reviewer,
      kind: type,
      occupiedParaId: sequential.survivor.id,
    });
    // The merge's properties fold into the carried pending change.
    expect(
      apply({
        reviewer: sequential.reviewer,
        story,
        operation: sequentialOperations[1],
      }).skipped,
    ).toEqual([]);
    const result = batch.reviewer.applyDocumentOperationsToStory({
      undefinedReferences: "refuse",
      story,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: operations(batch),
      },
    });
    expect(result.skipped).toEqual([]);
    await capacity({ reviewer: batch.reviewer, kind: type });
    batch.reviewer.acceptAll();
    sequential.reviewer.acceptAll();
    const content = (reviewer: FolioDocxReviewer) =>
      reviewer.snapshotStory(story)?.blocks.map((snapshotBlock) => {
        const paragraph = expectParagraphBlock(snapshotBlock);
        return {
          text: paragraph.text,
          directAlignment: paragraph.directAlignment,
        };
      });
    expect(content(batch.reviewer)).toEqual(content(sequential.reviewer));
  },
);
