/**
 * Comparison staging on a canonical headless reviewer gives the legacy
 * reviewer's package: the target's styles, numbering, final section change
 * and an imported header part land in the session's document as the legacy
 * reviewer stages them, each as a session commit, and survive a save.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "./headless";

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
});
