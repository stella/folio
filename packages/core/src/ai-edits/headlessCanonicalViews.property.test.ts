/**
 * A canonical headless reviewer's resolved views read as the legacy
 * reviewer's.
 *
 * For every story of a document, the original and final views (text,
 * blocks and the changes left) and the story's tables in each view are
 * read on both engines. A canonical reviewer reads them from a second
 * session resolved to the view, and leaves its own document as it was.
 * A story resolved to its final view keeps that view and holds no change.
 *
 * A note is part of a resolved view exactly when a reference to it survives
 * in that view, wherever the reference sits: in a run, a hyperlink, a content
 * control, a field or a table cell. A note whose every reference the view
 * removes leaves with it, as resolving the document removes it. A note no
 * reference reaches in the document itself stays in both views: resolving
 * removes only what a resolved change takes away.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { reviewerDocumentArbitrary } from "../../../../test/generators/reviewerDocumentArbitraries";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type {
  BlockContent,
  Document,
  Hyperlink,
  InlineSdt,
  Paragraph,
  ParagraphContent,
  Run,
  SimpleField,
} from "../types/document";
import {
  FOLIO_RESOLVED_REVIEWED_VIEWS,
  FolioDocxReviewer,
  type FolioResolvedReviewedView,
} from "./headless";

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

const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), { session });

/** Every story in both resolved views. */
const views = (reviewer: FolioDocxReviewer) =>
  reviewer.listStories().flatMap(({ handle }) =>
    FOLIO_RESOLVED_REVIEWED_VIEWS.map((view) => {
      const read = reviewer.readReviewedStory({ story: handle, view });
      return {
        handle,
        view,
        text: read?.text ?? null,
        changes: read?.changes ?? null,
        blocks: read?.snapshot.blocks ?? null,
        tables: reviewer
          .storyTables({ story: handle, view })
          .map((table) => table.node.textContent),
      };
    }),
  );

/** Every note a reference reaches anywhere in `value`, at any depth. */
const referencedNotes = (
  value: unknown,
  found: Set<string> = new Set(),
): Set<string> => {
  if (Array.isArray(value)) {
    for (const item of value) referencedNotes(item, found);
    return found;
  }
  if (typeof value !== "object" || value === null) return found;
  if (
    "type" in value &&
    "id" in value &&
    (value.type === "footnoteRef" || value.type === "endnoteRef") &&
    typeof value.id === "number"
  )
    found.add(
      `${value.type === "footnoteRef" ? "footnote" : "endnote"}:${String(value.id)}`,
    );
  for (const child of Object.values(value)) referencedNotes(child, found);
  return found;
};

/**
 * LEGACY_JOINED_TEXT_TAKES_MARK_RUN_PROPERTIES, which goes with the legacy
 * reviewer: a view that resolves a paragraph's mark (an inserted one in the
 * original view, a deleted one in the final view) joins its text into the
 * next paragraph whose mark stays. A mark's run properties format the mark
 * alone, so the joined text keeps its own run formatting and otherwise takes
 * the surviving paragraph's style; the legacy reviewer gives it the surviving
 * mark's run properties. For a surviving paragraph whose mark carries run
 * properties, the canonical preview runs are expected.
 */
const joinTargetsWithMarkRunProperties = (
  content: readonly BlockContent[],
  view: FolioResolvedReviewedView,
): Set<string> => {
  const resolved = view === "original" ? "ins" : "del";
  const targets = new Set<string>();
  const joins = (block: BlockContent | undefined): boolean =>
    block?.type === "paragraph" && block.pPrMark?.kind === resolved;
  content.forEach((block, index) => {
    if (!joins(block)) return;
    let next = index + 1;
    while (joins(content[next])) next += 1;
    const target = content[next];
    if (
      target?.type === "paragraph" &&
      target.formatting?.runProperties !== undefined &&
      target.paraId !== undefined
    )
      targets.add(target.paraId);
  });
  return targets;
};

const storyContent = (
  document: Document,
  handle: ReturnType<typeof views>[number]["handle"],
): readonly BlockContent[] => {
  const pkg = document.package;
  switch (handle.type) {
    case "main":
      return pkg.document.content;
    case "header":
      return pkg.headers?.get(handle.relationshipId)?.content ?? [];
    case "footer":
      return pkg.footers?.get(handle.relationshipId)?.content ?? [];
    case "footnote":
      return (
        pkg.footnotes?.find(({ id }) => id === handle.noteId)?.content ?? []
      );
    case "endnote":
      return (
        pkg.endnotes?.find(({ id }) => id === handle.noteId)?.content ?? []
      );
  }
};

/**
 * The legacy reviewer's views, less each note the view takes every reference
 * away from: the legacy reviewer reads a note's own state whether or not
 * anything still refers to it, while resolving the document removes it.
 */
const expectedViews = async (
  buffer: ArrayBuffer,
  canonical: ReturnType<typeof views>,
) => {
  const legacy = await open(buffer, "legacy");
  const source = legacy.toDocument();
  const referenced = referencedNotes(legacy.toDocument().package);
  const reached: Record<FolioResolvedReviewedView, Set<string>> = {
    original: referencedNotes(legacy.toResolvedDocument("reject").package),
    final: referencedNotes(legacy.toResolvedDocument("accept").package),
  };
  const expected = views(legacy);
  for (const read of expected) {
    const targets = joinTargetsWithMarkRunProperties(
      storyContent(source, read.handle),
      read.view,
    );
    if (targets.size === 0 || read.blocks === null) continue;
    const counterpart = canonical.find(
      ({ handle, view }) =>
        view === read.view &&
        JSON.stringify(handle) === JSON.stringify(read.handle),
    );
    for (const block of read.blocks) {
      if (!targets.has(block.id)) continue;
      const own = counterpart?.blocks?.find(({ id }) => id === block.id);
      if (own !== undefined && "previewRuns" in own && "previewRuns" in block)
        Object.assign(block, { previewRuns: own.previewRuns });
    }
  }
  for (const read of expected)
    if (
      (read.handle.type === "footnote" || read.handle.type === "endnote") &&
      referenced.has(`${read.handle.type}:${String(read.handle.noteId)}`) &&
      !reached[read.view].has(
        `${read.handle.type}:${String(read.handle.noteId)}`,
      )
    )
      Object.assign(read, {
        text: null,
        changes: null,
        blocks: null,
        tables: [],
      });
  return expected;
};

type ViewOutcome = "read" | "sessionRefused" | "resolutionRefused";

const viewLaws = async (buffer: ArrayBuffer): Promise<ViewOutcome> => {
  let canonical: FolioDocxReviewer;
  try {
    canonical = await open(buffer, "canonical");
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "sessionRefused";
    throw error;
  }
  const before = canonical.toDocument();
  let read: ReturnType<typeof views>;
  try {
    read = views(canonical);
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "resolutionRefused";
    throw error;
  }
  expect(read).toEqual(await expectedViews(buffer, read));
  // Reading a view resolves nothing in the reviewer itself.
  expect(canonical.toDocument()).toEqual(before);

  // Resolving one story reads, afterwards, as that story's view did.
  for (const { handle } of canonical.listStories()) {
    const reviewer = await open(buffer, "canonical");
    const final = reviewer.readReviewedStory({ story: handle, view: "final" });
    expect(
      reviewer.resolveReviewedStory({ story: handle, view: "final" }),
    ).toBe(true);
    // Its final view is unchanged, and no change of it is left pending.
    const resolved = reviewer.readReviewedStory({
      story: handle,
      view: "final",
    });
    expect(resolved?.text).toBe(final?.text);
    expect(
      reviewer.readReviewedStory({ story: handle, view: "current-markup" })
        ?.changes,
    ).toEqual([]);
  }
  return "read";
};

const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch {
    return null;
  }
};

describe("a canonical reviewer's resolved views", () => {
  test.each(FIXTURE_FILES)(
    "%s reads its resolved views as the legacy reviewer does",
    async (filename) => {
      await viewLaws(await ingested(readFixture(filename)));
    },
    propertyTestTimeout(30_000),
  );

  test(
    "generated documents read their resolved views as the legacy reviewer does",
    async () => {
      await assertProperty(
        fc.asyncProperty(reviewerDocumentArbitrary, async (document) => {
          const buffer = await packaged(document);
          if (buffer !== null) await viewLaws(buffer);
        }),
        {
          id: "generated documents read their resolved views as the legacy reviewer does",
          numRuns: 25,
        },
      );
    },
    propertyTestTimeout(120_000),
  );
});

const NOTE_ID = 37;
const NOTE = { type: "footnote", noteId: NOTE_ID } as const;
const REFERENCE: Run = {
  type: "run",
  content: [{ type: "footnoteRef", id: NOTE_ID }],
};

type Placement = "run" | "hyperlink" | "contentControl" | "field" | "tableCell";

/** Where the note's only reference sits. */
const PLACEMENTS: Record<Placement, Run | Hyperlink | InlineSdt | SimpleField> =
  {
    run: REFERENCE,
    hyperlink: {
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: [REFERENCE],
    } satisfies Hyperlink,
    contentControl: {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: 42, tag: "clause" },
      content: [REFERENCE],
    } satisfies InlineSdt,
    field: {
      type: "simpleField",
      instruction: "PAGE",
      fieldType: "PAGE",
      content: [REFERENCE],
    } satisfies SimpleField,
    tableCell: REFERENCE,
  };

/** How the reference is tracked, and the views it survives in. */
const TRACKING = {
  none: ["original", "final"],
  insertion: ["final"],
  deletion: ["original"],
} as const satisfies Record<string, readonly FolioResolvedReviewedView[]>;

const referenceParagraph = (
  placement: Placement,
  tracking: keyof typeof TRACKING,
): Paragraph => {
  const carrier = PLACEMENTS[placement];
  const info = { id: 91, author: "A" };
  if (tracking === "none") return paragraph([run("See "), carrier]);
  const content: ParagraphContent =
    tracking === "insertion"
      ? { type: "insertion", info, content: [carrier] }
      : { type: "deletion", info, content: [carrier] };
  return paragraph([run("See "), content]);
};

const placedDocument = (
  placement: Placement,
  tracking: keyof typeof TRACKING,
): Document => {
  const document = fromMarkdown("Body.");
  const host = referenceParagraph(placement, tracking);
  document.package.document.content.push(
    placement === "tableCell"
      ? {
          type: "table",
          columnWidths: [4800],
          rows: [
            {
              type: "tableRow",
              cells: [{ type: "tableCell", content: [host] }],
            },
          ],
        }
      : host,
    paragraph("After."),
  );
  document.package.footnotes = [
    { type: "footnote", id: NOTE_ID, content: [paragraph("Note text.")] },
  ];
  return document;
};

describe("a note in a canonical reviewer's resolved views", () => {
  for (const placement of Object.keys(PLACEMENTS) as Placement[]) {
    for (const tracking of Object.keys(TRACKING) as Array<
      keyof typeof TRACKING
    >) {
      test(
        `a reference in a ${placement}, ${tracking}: the note reads in each view its reference survives`,
        async () => {
          const buffer = await ingested(
            await createDocx(placedDocument(placement, tracking)),
          );
          const reviewer = await open(buffer, "canonical");
          const survives: readonly FolioResolvedReviewedView[] =
            TRACKING[tracking];
          for (const view of FOLIO_RESOLVED_REVIEWED_VIEWS) {
            const text =
              reviewer.readReviewedStory({ story: NOTE, view })?.text ?? null;
            if (survives.includes(view)) expect(text).toContain("Note text.");
            else expect(text).toBeNull();
          }
          // The reviewer's own document keeps the note whatever the views do.
          expect(
            reviewer.readReviewedStory({ story: NOTE, view: "current-markup" })
              ?.text,
          ).toContain("Note text.");
        },
        propertyTestTimeout(30_000),
      );
    }
  }
});

describe("text a resolved paragraph mark joins into another paragraph", () => {
  test("keeps its own run formatting, not the surviving mark's", async () => {
    const inserted = (id: number) => ({
      kind: "ins" as const,
      info: { id, author: "A" },
    });
    const document = fromMarkdown("x");
    document.package.document.content = [
      { type: "paragraph", content: [], pPrMark: inserted(1) },
      {
        type: "paragraph",
        content: [run("a")],
        pPrMark: inserted(2),
      },
      {
        type: "paragraph",
        content: [],
        // The mark alone is 8 pt; the joined text is not.
        formatting: { styleId: "Heading1", runProperties: { fontSize: 16 } },
      },
      // The same style with no mark formatting: the size its text takes.
      {
        type: "paragraph",
        content: [run("b")],
        formatting: { styleId: "Heading1" },
      },
    ];
    const reviewer = await open(
      await ingested(await createDocx(document)),
      "canonical",
    );
    const original = reviewer.readReviewedStory({
      story: { type: "main" },
      view: "original",
    });
    const sizeOf = (text: string) => {
      const block = original?.snapshot.blocks.find(
        (candidate) =>
          "previewRuns" in candidate &&
          candidate.previewRuns?.[0]?.text === text,
      );
      return block !== undefined && "previewRuns" in block
        ? block.previewRuns?.[0]?.fontSizePt
        : undefined;
    };
    expect(sizeOf("b")).toBeDefined();
    expect(sizeOf("b")).not.toBe(8);
    expect(sizeOf("a")).toBe(sizeOf("b"));
  });
});
