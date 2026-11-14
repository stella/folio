/**
 * Comparison staging on a canonical headless reviewer gives the legacy
 * reviewer's package: the target's styles, numbering, final section change,
 * section endpoint changes, an imported header part and a terminal table's
 * review carrier land in the session's document as the legacy reviewer
 * stages them, each as a session commit, and survive a save.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

import { buildBodySequenceDocx } from "../compare/__fixtures__/body-sequence";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Document, SectionProperties } from "../types/document";
import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "./headless";
import { sourceDocumentOf } from "./snapshot";

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
    expect(endpoints(reopened)).toEqual(endpoints(canonical.base));
  });

  test("a terminal table's review carrier stages as on the legacy reviewer", async () => {
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
    const canonical = await staged("canonical");
    expect(legacy.result).toBe(true);
    expect(canonical.result).toBe(legacy.result);
    // The carrier's paragraph id is minted by each engine; the rest agrees.
    const body = (reviewer: FolioDocxReviewer) =>
      reviewer.toDocument().package.document.content.map((block) =>
        block.type === "paragraph"
          ? {
              type: block.type,
              reviewCarrier: block.reviewCarrier,
              content: block.content,
            }
          : { type: block.type },
      );
    expect(body(canonical.base)).toEqual(body(legacy.base));
    expect(canonical.base.getContentAsText()).toBe(
      legacy.base.getContentAsText(),
    );
  });
});
