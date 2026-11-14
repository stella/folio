/**
 * The headless reviewer's reads agree on both engines.
 *
 * A canonical reviewer keeps its document in the canonical session: the body
 * is the session's projection and the model is the session's document. The
 * reads themselves are one implementation over either engine, so every read
 * must give what the legacy reviewer gives for the same package.
 *
 * Both engines read the package a host stores after `ensureParaIds`, which is
 * also what the canonical session normalizes its input to. On a package with
 * unidentified paragraphs the engines mint different ids (legacy in PM,
 * canonical in the package); that difference is stated in its own test
 * below rather than hidden behind the normalization.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import type { Node as PMNode } from "prosemirror-model";

import { reviewerDocumentArbitrary } from "../../../../test/generators/reviewerDocumentArbitraries";
import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  CanonicalSessionError,
  createCanonicalSession,
} from "../canonical/canonicalSession";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationBatch,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import {
  footnoteToProseDoc,
  headerFooterToProseDoc,
  toProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import {
  CANONICAL_GAP,
  type HeadlessReviewerSession,
} from "../types/canonicalCapabilities";
import type {
  BlockContent,
  Document,
  Endnote,
  Footnote,
  Paragraph,
} from "../types/document";
import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "./headless";
import { sourceDocumentOf } from "./snapshot";
import type { FolioAIEditSnapshot } from "./types";

const FIXTURES_DIR = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus",
);

const FIXTURE_FILES: string[] = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
};

/** The package as a host stores it: every paragraph and control identified. */
const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

type NodeJson = {
  type: string;
  attrs?: Record<string, unknown>;
  content?: NodeJson[];
};

const storyJson = (node: PMNode): NodeJson => node.toJSON();

/**
 * A note as converted, its automatic mark (a leading `noteMarker` atom) told
 * apart from the rest so the one stated difference can be expected exactly.
 */
const noteView = (node: PMNode) => {
  const json = storyJson(node);
  const [first, ...rest] = json.content ?? [];
  const mark = first?.content?.at(0)?.type === "noteMarker";
  return {
    automaticMark: mark,
    content:
      mark && first
        ? [{ ...first, content: first.content?.slice(1) }, ...rest]
        : json.content,
  };
};

/**
 * The model as the shared readers see it: each story converted by the same
 * converter, the comment threads, the numbering and the final section.
 */
const modelView = (document: Document) => {
  const options = {
    ...(document.package.styles !== undefined && {
      styles: document.package.styles,
    }),
    ...(document.package.numbering !== undefined && {
      numbering: document.package.numbering,
    }),
    ...(document.package.theme !== undefined && {
      theme: document.package.theme,
    }),
  };
  const story = (content: BlockContent[]) =>
    storyJson(headerFooterToProseDoc(content, options));
  const note = ({ id, noteType, content }: Footnote | Endnote) => ({
    id,
    noteType,
    ...noteView(footnoteToProseDoc(content, options)),
  });
  return {
    body: storyJson(toProseDoc(document)),
    headers: [...(document.package.headers ?? [])].map(([id, header]) => [
      id,
      header.hdrFtrType,
      story(header.content),
    ]),
    footers: [...(document.package.footers ?? [])].map(([id, footer]) => [
      id,
      footer.hdrFtrType,
      story(footer.content),
    ]),
    footnotes: (document.package.footnotes ?? []).map(note),
    endnotes: (document.package.endnotes ?? []).map(note),
    comments: (document.package.document.comments ?? []).map(
      ({ id, parentId, author, date, done }) => ({
        id,
        parentId,
        author,
        date,
        done,
      }),
    ),
    numbering: document.package.numbering ?? null,
    finalSection: document.package.document.finalSectionProperties ?? null,
  };
};

/** Whether a story's first paragraph starts with an automatic note mark. */
const leadingNoteMark = (snapshot: FolioAIEditSnapshot | null): boolean =>
  snapshot !== null &&
  sourceDocumentOf(snapshot).firstChild?.firstChild?.type.name === "noteMarker";

/** Every read the reviewer offers, as plain data. */
const reads = (reviewer: FolioDocxReviewer) => {
  const snapshot = reviewer.snapshot();
  const { blocks, anchors } = snapshot;
  const stories = reviewer.listStories();
  const storySnapshots = stories.map(({ handle }) =>
    reviewer.snapshotStory(handle),
  );
  return {
    snapshot: { blocks, anchors },
    // The body the snapshot was taken from: on canonical, the projection.
    body: storyJson(sourceDocumentOf(snapshot)),
    content: reviewer.getContent(),
    text: reviewer.getContentAsText(),
    annotated: reviewer.getContentAsText({ annotated: true }),
    notes: reviewer.getNotesAsText(),
    stories,
    readStory: stories.map(({ handle }) => reviewer.readStory(handle)),
    snapshotStory: storySnapshots.map(
      (story) => story && { blocks: story.blocks, anchors: story.anchors },
    ),
    noteMarks: storySnapshots.map(leadingNoteMark),
    readReviewedStory: stories.map(({ handle }) => {
      const read = reviewer.readReviewedStory({
        story: handle,
        view: "current-markup",
      });
      return (
        read && {
          text: read.text,
          changes: read.changes,
          blocks: read.snapshot.blocks,
        }
      );
    }),
    storyTables: reviewer
      .storyTables({ view: "current-markup" })
      .map((table) => storyJson(table.node)),
    changes: reviewer.getChanges(),
    comments: reviewer.getComments(),
    numbering: reviewer.readNumberingDefinitions(),
    properties: reviewer.getDocumentProperties(),
    pendingSuggestions: reviewer.exportPendingSuggestions(),
    document: modelView(reviewer.toDocument()),
  };
};

type ReviewerReads = ReturnType<typeof reads>;
type StoryAnchors = FolioAIEditSnapshot["anchors"];

/** Anchors with one atom inserted at the start of the story's first block. */
const pastLeadingAtom = (anchors: StoryAnchors): StoryAnchors => {
  const first = Math.min(...Object.values(anchors).map(({ from }) => from));
  return Object.fromEntries(
    Object.entries(anchors).map(([key, anchor]) => [
      key,
      {
        ...anchor,
        from: anchor.from === first ? anchor.from : anchor.from + 1,
        to: anchor.to + 1,
      },
    ]),
  );
};

/** A converted story with the ids legacy mints for paragraphs only PM has. */
const withoutPositionalIds = (node: NodeJson): NodeJson => ({
  ...node,
  ...(node.attrs?.["idStability"] === "positional" && {
    attrs: { ...node.attrs, paraId: null, idStability: undefined },
  }),
  ...(node.content !== undefined && {
    content: node.content.map(withoutPositionalIds),
  }),
});

/**
 * What the canonical engine must read, given what the legacy engine read.
 *
 * One difference is the canonical one by design: the session's model owns a
 * normal note's automatic mark (`@stll/docx-core` ops, `noteMarks.ts`), so a
 * footnote or endnote whose package spells no mark reads with one, the atom
 * before its text, as a canonical editor shows it, and the model
 * `toDocument` returns holds it. Only that note's anchor positions move; its
 * text and block ids do not.
 *
 * Two are legacy's. `toDocument` on legacy is its PM save projection of the
 * package, which respells what it carries (a bookmark around a paragraph as
 * inline boundaries, a paragraph's run defaults, the cached rendered page break
 * hints); the canonical model is the session's, which gives the package back
 * as parsed. So the canonical model is held to the package itself, read by
 * the same view.
 *
 * And legacy gives a paragraph the package does not hold (the
 * filler of a table cell the grid omits) a positional id, which the package
 * never carries. The input is ingested, so a positional id can only be one of
 * those; canonical leaves them unidentified, as the projection has them.
 */
/** Whether `value` holds, at any depth, a record `matches` accepts. */
const holdsRecord = (
  value: unknown,
  matches: (record: Record<string, unknown>) => boolean,
): boolean => {
  if (value instanceof Map)
    return [...value.values()].some((entry) => holdsRecord(entry, matches));
  if (Array.isArray(value))
    return value.some((entry) => holdsRecord(entry, matches));
  if (
    typeof value !== "object" ||
    value === null ||
    value instanceof ArrayBuffer
  )
    return false;
  const record = Object.fromEntries(Object.entries(value));
  return (
    matches(record) ||
    Object.values(record).some((entry) => holdsRecord(entry, matches))
  );
};

type NoteHandle = { type: "footnote" | "endnote"; noteId: number };

/**
 * Whether the session's model gives a note its automatic mark: a normal note
 * opening with a paragraph, spelling no mark, that no reference marks with
 * custom text.
 */
const sessionMarksNote = (parsed: Document, { type, noteId }: NoteHandle) => {
  const notes =
    type === "footnote" ? parsed.package.footnotes : parsed.package.endnotes;
  const note = notes?.find(({ id }) => id === noteId);
  if (!note || (note.noteType !== undefined && note.noteType !== "normal"))
    return false;
  if (note.content.at(0)?.type !== "paragraph") return false;
  if (holdsRecord(note.content, (record) => record["type"] === "noteMarker"))
    return false;
  const reference = type === "footnote" ? "footnoteRef" : "endnoteRef";
  return !holdsRecord(
    parsed.package,
    (record) =>
      record["type"] === reference &&
      record["id"] === noteId &&
      record["customMarkFollows"] === true,
  );
};

type ExpectationOptions = {
  legacy: ReviewerReads;
  /** The stored package's model, as the parser reads it. */
  parsed: Document;
};

const canonicalExpectation = ({
  legacy,
  parsed,
}: ExpectationOptions): ReviewerReads => {
  const markedByTheSession = legacy.stories.map(
    ({ handle }, index) =>
      (handle.type === "footnote" || handle.type === "endnote") &&
      !legacy.noteMarks[index] &&
      sessionMarksNote(parsed, handle),
  );
  const packageModel = modelView(parsed);
  const notes = (
    type: NoteHandle["type"],
    packageNotes: ReviewerReads["document"]["footnotes"],
  ) =>
    packageNotes.map((note) => ({
      ...note,
      automaticMark:
        note.automaticMark ||
        sessionMarksNote(parsed, { type, noteId: note.id }),
    }));
  return {
    ...legacy,
    body: withoutPositionalIds(legacy.body),
    storyTables: legacy.storyTables.map(withoutPositionalIds),
    document: {
      ...packageModel,
      footnotes: notes("footnote", packageModel.footnotes),
      endnotes: notes("endnote", packageModel.endnotes),
    },
    snapshotStory: legacy.snapshotStory.map((snapshot, index) =>
      snapshot && markedByTheSession[index]
        ? {
            blocks: snapshot.blocks,
            anchors: pastLeadingAtom(snapshot.anchors),
          }
        : snapshot,
    ),
    noteMarks: legacy.noteMarks.map(
      (mark, index) => mark || markedByTheSession[index] === true,
    ),
  };
};

type EngineOutcome =
  { type: "read"; reads: ReviewerReads } | { type: "refused"; error: unknown };

const openOn = async (
  buffer: ArrayBuffer,
  session: HeadlessReviewerSession,
): Promise<EngineOutcome> => {
  try {
    const reviewer = await FolioDocxReviewer.fromBuffer(buffer.slice(0), {
      session,
    });
    return { type: "read", reads: reads(reviewer) };
  } catch (error) {
    return { type: "refused", error };
  }
};

/** The canonical editor's own verdict on the package: the session seed. */
const sessionRefusal = async (
  buffer: ArrayBuffer,
): Promise<CanonicalSessionError | null> => {
  const document = await parseDocx(buffer.slice(0), {
    detectVariables: false,
    preloadFonts: false,
  });
  const created = createCanonicalSession(document, document.package.styles);
  return created.isErr() ? created.error : null;
};

type ReadParity = "read" | "unreadable" | "sessionRefused";

/**
 * Open the stored package on both engines. Both read and agree; or neither
 * opens it; or the session itself refuses it, with the typed refusal a
 * canonical editor gets for the same package.
 */
const compareEngines = async (stored: ArrayBuffer): Promise<ReadParity> => {
  const buffer = await ingested(stored);
  const legacy = await openOn(buffer, "legacy");
  const canonical = await openOn(buffer, "canonical");
  if (legacy.type === "refused") {
    expect(canonical.type).toBe("refused");
    return "unreadable";
  }
  if (canonical.type === "refused") {
    const verdict = await sessionRefusal(buffer);
    expect(CanonicalSessionError.is(canonical.error)).toBe(true);
    expect(verdict).not.toBeNull();
    expect(canonical.error).toMatchObject({
      message: verdict?.message,
      ground: verdict?.ground,
    });
    return "sessionRefused";
  }
  const expected = canonicalExpectation({
    legacy: legacy.reads,
    parsed: await parseDocx(buffer.slice(0), {
      detectVariables: false,
      preloadFonts: false,
    }),
  });
  const differing = (Object.keys(expected) as (keyof ReviewerReads)[]).filter(
    (read) => !Bun.deepEquals(canonical.reads[read], expected[read]),
  );
  // Name the reads that differ before showing the first one in full.
  const first = differing.at(0);
  expect({
    differing,
    first: first && canonical.reads[first],
  }).toEqual({ differing: [], first: first && expected[first] });
  return "read";
};

type GeneratedParity = ReadParity | "notAPackage";

/** The generated model as a package, or nothing when the writer refuses it. */
const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await createDocx(structuredClone(document));
  } catch {
    return null;
  }
};

const generatedParity = (tally: Map<GeneratedParity, number>) =>
  fc.asyncProperty(reviewerDocumentArbitrary, async (document) => {
    const buffer = await packaged(document);
    const outcome =
      buffer === null ? "notAPackage" : await compareEngines(buffer);
    tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
  });

/** A numbered item inside a content control, as a group control holds it. */
const numberedItem = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  formatting: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
  content: [{ type: "run", content: [{ type: "text", text: "Item" }] }],
});

/**
 * A list that continues out of a content control, drawn on every run: the
 * item after the control is the list's second item.
 */
const listAcrossControl = (): Document => {
  const content: BlockContent[] = [
    {
      type: "blockSdt",
      properties: { sdtType: "group", id: 1 },
      content: [numberedItem("00000001")],
    },
    numberedItem("00000002"),
  ];
  const page = { pageWidth: 12240, pageHeight: 15840 };
  return {
    package: {
      document: {
        content,
        sections: [{ properties: page, content }],
        finalSectionProperties: page,
      },
      numbering: {
        abstractNums: [
          {
            abstractNumId: 1,
            levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
          },
        ],
        nums: [{ numId: 1, abstractNumId: 1 }],
      },
      settings: { defaultTabStop: 720 },
    },
  };
};

const GENERATED_RUNS = 30;
const COVERAGE_RUNS = 20;

describe("headless reviewer reads agree on both engines", () => {
  test("corpus fixtures are present", () => {
    expect(FIXTURE_FILES.length).toBeGreaterThan(0);
  });

  test.each(FIXTURE_FILES)(
    "%s reads the same on the canonical session",
    async (filename) => {
      expect(await compareEngines(readFixture(filename))).not.toBe(
        "unreadable",
      );
    },
    propertyTestTimeout(20_000),
  );

  test(
    "generated documents read the same on the canonical session",
    async () => {
      await assertProperty(generatedParity(new Map()), {
        id: "generated documents read the same on the canonical session",
        numRuns: GENERATED_RUNS,
        examples: [[listAcrossControl()]],
      });
    },
    propertyTestTimeout(60_000),
  );

  test(
    "most generated documents are read by both engines",
    async () => {
      const tally = new Map<GeneratedParity, number>();
      await checkCoverageSample(generatedParity(tally), COVERAGE_RUNS);
      expect(tally.get("read") ?? 0).toBeGreaterThanOrEqual(COVERAGE_RUNS / 2);
    },
    propertyTestTimeout(60_000),
  );

  test.each(FIXTURE_FILES)(
    "%s: unidentified paragraphs get the package's ids on canonical",
    async (filename) => {
      // Legacy mints missing paragraph ids in PM; canonical, like the
      // canonical editor, reads the package after `ensureParaIds`. Both are
      // deterministic per input; the canonical reading of the stored bytes is
      // the reading of the identified package.
      const buffer = readFixture(filename);
      expect(await openOn(buffer, "canonical")).toEqual(
        await openOn(await ingested(buffer), "canonical"),
      );
    },
    propertyTestTimeout(20_000),
  );
});

describe("a canonical reviewer refuses what is not on the session yet", () => {
  const refusal = {
    _tag: "CanonicalSessionError",
    reason: "refused",
    ground: { gap: CANONICAL_GAP.publicHeadlessSession },
  };
  const batch: FolioDocumentOperationBatch = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    operations: [],
    mode: "tracked-changes",
  };

  const expectRefusal = (action: () => unknown) => {
    let thrown: unknown;
    try {
      action();
    } catch (error) {
      thrown = error;
    }
    expect(CanonicalSessionError.is(thrown)).toBe(true);
    expect(thrown).toMatchObject(refusal);
  };

  test("every mutation, resolution, undo and save throws the typed refusal", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      readFixture("upstream-styled-content.docx"),
      { session: "canonical" },
    );
    const before = reads(reviewer);
    expectRefusal(() =>
      reviewer.applyOperations([], { undefinedReferences: "refuse" }),
    );
    expectRefusal(() =>
      reviewer.applyDocumentOperations(batch, {
        undefinedReferences: "refuse",
      }),
    );
    expectRefusal(() =>
      reviewer.applyDocumentOperationsToStory({
        story: { type: "main" },
        batch,
        undefinedReferences: "refuse",
      }),
    );
    expectRefusal(() =>
      reviewer.undoDocumentOperations({
        type: "documentOperationUndo",
        id: "undo",
      }),
    );
    expectRefusal(() => reviewer.loadPendingSuggestions([]));
    expectRefusal(() => reviewer.replyTo(1, { text: "reply" }));
    expectRefusal(() => reviewer.resolveComment("1"));
    expectRefusal(() => reviewer.acceptChange(1));
    expectRefusal(() => reviewer.rejectChange(1));
    expectRefusal(() => reviewer.acceptSuggestion("suggestion"));
    expectRefusal(() => reviewer.rejectSuggestion("suggestion"));
    expectRefusal(() => reviewer.acceptAll());
    expectRefusal(() => reviewer.rejectAll());
    expectRefusal(() => reviewer.toResolvedDocument("accept"));
    expectRefusal(() => reviewer.resolveReviewedStory({ view: "final" }));
    expectRefusal(() => reviewer.readReviewedStory({ view: "final" }));
    expectRefusal(() => reviewer.storyTables({ view: "original" }));
    expectRefusal(() =>
      reviewer.matchStoryTableGeometry({
        targetTables: new Map(),
        pairings: [],
        revisionStamp: { date: "2026-01-02T03:04:05Z", idSeed: 1 },
      }),
    );
    expectRefusal(() => getFolioDocxComparisonAccess(reviewer));
    await expect(reviewer.toBuffer()).rejects.toMatchObject(refusal);
    await expect(reviewer.save()).rejects.toMatchObject(refusal);
    // A refusal changes nothing a read sees.
    expect(reads(reviewer)).toEqual(before);
  });
});
