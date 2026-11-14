/**
 * Comparison staging on a canonical headless reviewer gives the legacy
 * reviewer's package: the target's styles, numbering, final section change,
 * section endpoint changes, an imported header part and a paired table's own
 * properties land in the session's document as the legacy reviewer stages
 * them, each as a session commit, and survive a save. A terminal table's
 * review carrier refuses under a typed gap.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

import { CanonicalSessionError } from "../canonical/canonicalSession";
import { buildBodySequenceDocx } from "../compare/__fixtures__/body-sequence";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import {
  CANONICAL_GAP,
  type HeadlessReviewerSession,
} from "../types/canonicalCapabilities";
import type { Document, SectionProperties } from "../types/document";
import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "./headless";
import { sourceDocumentOf } from "./snapshot";
import { projectTableGeometry } from "./table-geometry";

const corpus = (name: string) =>
  path.join(import.meta.dir, "../docx/__tests__/__fixtures__/corpus", name);

const ingested = async (name: string): Promise<ArrayBuffer> => {
  const bytes = readFileSync(corpus(name));
  const prepared = await prepareCanonicalDocxInput(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const BASE = "upstream-styled-content.docx";
const TARGET_STYLES = "upstream-complex-styles.docx";
const TARGET_HEADERS = "step3-header-footer-fields.docx";

const open = async (name: string, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(await ingested(name), {
    author: "Comparison",
    session,
  });

const openBuffer = async (
  buffer: ArrayBuffer,
  session: HeadlessReviewerSession,
) => {
  const prepared = await prepareCanonicalDocxInput(buffer.slice(0));
  if (prepared.isErr()) throw prepared.error;
  return FolioDocxReviewer.fromBuffer(prepared.value.slice().buffer, {
    author: "Comparison",
    session,
  });
};

const PORTRAIT: SectionProperties = { pageWidth: 12240, pageHeight: 15840 };
const LANDSCAPE: SectionProperties = {
  pageWidth: 15840,
  pageHeight: 12240,
  orientation: "landscape",
};

/** Two sections: the first ends at "Before" with `first` as its properties. */
const twoSections = (first: SectionProperties): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            { type: "run", content: [{ type: "text", text: "Before" }] },
          ],
          sectionProperties: first,
        },
        {
          type: "paragraph",
          paraId: "00000002",
          content: [
            { type: "run", content: [{ type: "text", text: "After" }] },
          ],
        },
      ],
      finalSectionProperties: PORTRAIT,
    },
  },
});

/** One table of two cells, its own, its row's and its first cell's properties given. */
const oneTable = ({
  fill,
  height,
  justification,
}: {
  fill: string;
  height: number;
  justification: "left" | "center";
}) =>
  buildBodySequenceDocx([
    { kind: "paragraph", text: "Schedule" },
    {
      kind: "table",
      properties: { justification },
      rows: [
        {
          height,
          cells: [{ content: "Amount", shadingFill: fill }, "Due"],
        },
      ],
    },
    { kind: "paragraph", text: "End" },
  ]);

/** A body that ends in a table, with no paragraph after it. */
const endingInATable = async (): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(
    await buildBodySequenceDocx([
      { kind: "paragraph", text: "The source clause changes." },
      { kind: "table", rows: [["Source", "Amount"]] },
    ]),
  );
  const entry = zip.entry("word/document.xml");
  if (!entry) throw new Error("The fixture has no main document part.");
  const xml = await entry.text();
  const withoutCarrier = xml.replace(/<w:p\b[^>]*><\/w:p>(?=<w:sectPr>)/u, "");
  if (withoutCarrier === xml)
    throw new Error("The fixture has no terminal empty body paragraph.");
  zip.set("word/document.xml", withoutCarrier);
  return (await zip.generate({ compression: "store" })).buffer;
};

const REVISION = { date: "2026-01-02T03:04:05Z", idSeed: 9000 } as const;

/** Stage the target's styles, numbering and final section into the base. */
const stage = async (session: HeadlessReviewerSession) => {
  const base = await open(BASE, session);
  const target = await open(TARGET_STYLES, session);
  const access = getFolioDocxComparisonAccess(base);
  const targetAccess = getFolioDocxComparisonAccess(target);
  const snapshots = target
    .listStories()
    .flatMap(({ handle }) => target.snapshotStory(handle) ?? []);
  const styles = access.stageTargetStyles({
    source: target,
    snapshots,
    importedHeaderFooterSnapshots: [],
  });
  const previous = access.finalSectionProperties();
  const next = targetAccess.finalSectionProperties();
  const section =
    previous !== undefined && next !== undefined
      ? access.stageFinalSectionProperties({
          target: next,
          previous,
          revision: REVISION,
        })
      : false;
  return { base, styles, section };
};

describe("comparison staging on the canonical session", () => {
  test("the target's styles and final section stage as on the legacy reviewer", async () => {
    const legacy = await stage("legacy");
    const canonical = await stage("canonical");
    expect(canonical.styles.status).toBe(legacy.styles.status);
    expect(canonical.section).toBe(legacy.section);
    const packageOf = (reviewer: FolioDocxReviewer) => {
      const { styles, numbering, document } = reviewer.toDocument().package;
      return {
        styles,
        numbering,
        finalSection: document.finalSectionProperties,
      };
    };
    expect(packageOf(canonical.base)).toEqual(packageOf(legacy.base));
    expect(canonical.base.getChanges()).toEqual(legacy.base.getChanges());
    // The staged package saves and reopens as it was staged.
    const reopened = await FolioDocxReviewer.fromBuffer(
      await canonical.base.toBuffer(),
      { session: "canonical" },
    );
    expect(packageOf(reopened).styles).toEqual(
      packageOf(canonical.base).styles,
    );
    expect(reopened.getChanges()).toEqual(canonical.base.getChanges());
  });

  test("an imported header part joins the package unreferenced", async () => {
    const imported = async (session: HeadlessReviewerSession) => {
      const base = await open(BASE, session);
      const target = await open(TARGET_HEADERS, session);
      const header = target
        .listStories()
        .find(({ handle }) => handle.type === "header")?.handle;
      if (header?.type !== "header")
        throw new Error("The target fixture has no header.");
      const handle = await getFolioDocxComparisonAccess(
        base,
      ).createComparisonHeaderFooter(target, header);
      return { base, handle };
    };
    const legacy = await imported("legacy");
    const canonical = await imported("canonical");
    expect(canonical.handle).toEqual(legacy.handle);
    if (canonical.handle === null) return;
    expect(canonical.base.readStory(canonical.handle)).toEqual(
      legacy.base.readStory(canonical.handle),
    );
    expect(
      canonical.base.toDocument().package.document.finalSectionProperties,
    ).toEqual(legacy.base.toDocument().package.document.finalSectionProperties);
  });

  test("a section endpoint change stages as on the legacy reviewer", async () => {
    const staged = async (session: HeadlessReviewerSession) => {
      const base = await openBuffer(
        await createDocx(twoSections(PORTRAIT)),
        session,
      );
      const target = await openBuffer(
        await createDocx(twoSections(LANDSCAPE)),
        session,
      );
      const snapshot = target.snapshotStory({ type: "main" });
      if (!snapshot) throw new Error("The target has no main story.");
      const result = getFolioDocxComparisonAccess(
        base,
      ).stageMappedSectionBoundaries({
        target: sourceDocumentOf(snapshot),
        originalRevisionIdSeed: REVISION.idSeed,
        maxRanges: 10,
        revision: REVISION,
        mapTargetProperties: ({ kind, current, target: properties }) => {
          if (kind === "inserted") return { kind, target: properties };
          if (current === undefined) return null;
          return { kind, previous: current, target: properties };
        },
      });
      return { base, result };
    };
    const legacy = await staged("legacy");
    const canonical = await staged("canonical");
    expect(legacy.result).toMatchObject({
      status: "matched",
      rangeCount: 1,
      documentChanged: true,
    });
    expect(canonical.result).toEqual(legacy.result);
    const endpoints = (reviewer: FolioDocxReviewer) =>
      reviewer
        .toDocument()
        .package.document.content.flatMap((block) =>
          block.type === "paragraph" && block.sectionProperties !== undefined
            ? [block.sectionProperties]
            : [],
        );
    expect(endpoints(canonical.base)).toEqual(endpoints(legacy.base));
    expect(canonical.base.getChanges()).toEqual(legacy.base.getChanges());
    const reopened = await FolioDocxReviewer.fromBuffer(
      await canonical.base.toBuffer(),
      { session: "canonical" },
    );
    // `currentProperties` is the model's own copy of the new set; the
    // package holds only the previous one.
    const saved = (properties: readonly SectionProperties[]) =>
      properties.map(({ propertyChanges, ...rest }) => ({
        ...rest,
        propertyChanges: propertyChanges?.map(
          ({ currentProperties: _current, ...change }) => change,
        ),
      }));
    expect(saved(endpoints(reopened))).toEqual(
      saved(endpoints(canonical.base)),
    );
  });

  test("a terminal table's review carrier refuses on the session", async () => {
    const target = await openBuffer(
      await buildBodySequenceDocx([
        { kind: "paragraph", text: "The target clause changes." },
        { kind: "paragraph", text: "Terminal target paragraph." },
      ]),
      "legacy",
    );
    const snapshot = target.snapshotStory({ type: "main" });
    if (!snapshot) throw new Error("The target has no main story.");
    const staged = async (session: HeadlessReviewerSession) => {
      const base = await openBuffer(await endingInATable(), session);
      const result = getFolioDocxComparisonAccess(
        base,
      ).stageTerminalTableReviewCarrier(sourceDocumentOf(snapshot));
      return { base, result };
    };
    const legacy = await staged("legacy");
    expect(legacy.result).toBe(true);
    // A canonical session holds no review-carrier paragraph yet: staging one
    // refuses under the story content gap and leaves the document as it was.
    const base = await openBuffer(await endingInATable(), "canonical");
    const before = base.toDocument();
    let thrown: unknown;
    try {
      getFolioDocxComparisonAccess(base).stageTerminalTableReviewCarrier(
        sourceDocumentOf(snapshot),
      );
    } catch (error) {
      thrown = error;
    }
    expect(CanonicalSessionError.is(thrown)).toBe(true);
    expect(thrown).toMatchObject({
      reason: "refused",
      ground: { gap: CANONICAL_GAP.storyContentProjection },
    });
    expect(base.toDocument()).toEqual(before);
  });

  test("a paired table's own properties match as on the legacy reviewer", async () => {
    const matched = async (session: HeadlessReviewerSession) => {
      const base = await openBuffer(
        await oneTable({ fill: "FFFFFF", height: 400, justification: "left" }),
        session,
      );
      const target = await openBuffer(
        await oneTable({
          fill: "FF0000",
          height: 800,
          justification: "center",
        }),
        session,
      );
      const targetTables = new Map(
        target
          .storyTables({ view: "current-markup" })
          .map(({ index, node }) => [index, node] as const),
      );
      const next = base.matchStoryTableGeometry({
        targetTables,
        pairings: [0, 1].map((cellIndex) => ({
          base: { tableIndex: 0, rowIndex: 0, cellIndex },
          target: { tableIndex: 0, rowIndex: 0, cellIndex },
        })),
        revisionStamp: REVISION,
      });
      return { base, target, next };
    };
    const geometry = (
      reviewer: FolioDocxReviewer,
      view: "original" | "final" | "current-markup",
    ) => projectTableGeometry(reviewer.storyTables({ view }));
    const legacy = await matched("legacy");
    const canonical = await matched("canonical");
    expect(legacy.next).toBeGreaterThan(REVISION.idSeed);
    expect(canonical.next).toBe(legacy.next);
    expect(
      canonical.base.getChanges().map(({ type, author }) => ({ type, author })),
    ).toEqual(
      legacy.base.getChanges().map(({ type, author }) => ({ type, author })),
    );
    // Accepting gives the target's geometry; rejecting gives the base's.
    expect(geometry(canonical.base, "final")).toEqual(
      geometry(legacy.base, "final"),
    );
    expect(geometry(canonical.base, "final")).toEqual(
      geometry(canonical.target, "current-markup"),
    );
    expect(geometry(canonical.base, "original")).toEqual(
      geometry(legacy.base, "original"),
    );
    const reopened = await FolioDocxReviewer.fromBuffer(
      await canonical.base.toBuffer(),
      { session: "canonical" },
    );
    expect(geometry(reopened, "final")).toEqual(
      geometry(canonical.base, "final"),
    );
    expect(geometry(reopened, "original")).toEqual(
      geometry(canonical.base, "original"),
    );
  });
});
