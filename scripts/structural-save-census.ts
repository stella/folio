import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
/**
 * Which documents a structural selective save could splice, report-only.
 *
 * After Enter, or any edit that adds, removes or reorders paragraphs, the
 * selective save writes the body as a structural splice
 * (packages/core/src/docx/structuralXmlPatch.ts), which declines, and the save
 * falls back to the full repack, for reasons the source body alone decides:
 * a paragraph without a valid unique paraId, or a body block that is neither
 * a paragraph, a table nor a range marker. This scans each document's main
 * part for those reasons, mirroring the patch's static checks; the reasons
 * that depend on the edit (a table changed shape, the final section changed,
 * a range or field the edit would split) are not here. It also counts the
 * documents holding bookmarks, fields, permission or move ranges, which the
 * splice carries as long as the edit splits none of them.
 *
 * It also runs the real save once per document: paragraph ids stamped as a
 * host does at import, the document opened in a canonical session, Enter at
 * the end of its middle body paragraph, and a selective save. It records
 * whether that save spliced the body or, when it declined, why.
 *
 * Usage: bun scripts/structural-save-census.ts <out dir>
 * Scans the checked-in .docx files and, when it has been fetched, the pinned
 * public corpus; writes <out dir>/census.json and <out dir>/census.md.
 */
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../packages/core/src/canonical/canonicalSession";
import { serializeCanonicalSave } from "../packages/core/src/docx/canonicalSave";
import { ensureParaIds } from "../packages/core/src/docx/ensureParaIds";
import { parseDocx } from "../packages/core/src/docx/parser";
import type { SelectiveSaveAttempt } from "../packages/core/src/layout-engine/layoutInstrumentation";
import { schema } from "../packages/core/src/prosemirror/schema";
import {
  loadCorpusLock,
  REPOSITORY_ROOT,
  sourceCheckoutPath,
} from "./lib/corpus-manifest";

const WORDPROCESSINGML = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);

const RANGE_REASONS: Readonly<Record<string, string>> = {
  bookmarkStart: "bookmark",
  bookmarkEnd: "bookmark",
  permStart: "permission range",
  permEnd: "permission range",
  fldChar: "field character",
  moveFromRangeStart: "move range",
  moveFromRangeEnd: "move range",
  moveToRangeStart: "move range",
  moveToRangeEnd: "move range",
};
const BODY_BLOCKS = new Set([
  "p",
  "tbl",
  "sectPr",
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);
const PARA_ID = /\s[\w.-]+:paraId\s*=\s*["']([^"']*)["']/u;
const TAG = /<(\/?)([\w.-]+:)?([\w.-]+)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/gu;
const NAMESPACE = /\sxmlns(?::([\w.-]+))?\s*=\s*["']([^"']*)["']/gu;

/** What the real save did after one Enter. */
type Trial =
  | { outcome: "spliced" }
  | { outcome: "declined"; reason: string }
  | { outcome: "not run"; reason: string };

type Scan =
  | { readable: false; reason: string }
  | {
      readable: true;
      paragraphs: number;
      reasons: string[];
      ranges: string[];
      otherBodyBlocks: string[];
      trial?: Trial;
    };

/** The static reasons a structural splice of this main part declines. */
const scanMainPart = (xml: string): Scan => {
  const prefixes = new Map<string, string>();
  for (const [, prefix, uri] of xml.matchAll(NAMESPACE))
    if (!prefixes.has(prefix ?? "")) prefixes.set(prefix ?? "", uri ?? "");
  const reasons = new Set<string>();
  const ranges = new Set<string>();
  const otherBodyBlocks = new Set<string>();
  const paraIds = new Set<string>();
  let paragraphs = 0;
  let depth = 0;
  let bodyDepth: number | null = null;
  for (const match of xml.matchAll(TAG)) {
    const [, closing, prefix, local = "", attributes = "", selfClosing] = match;
    const uri = prefixes.get(prefix?.slice(0, -1) ?? "") ?? "";
    const word = WORDPROCESSINGML.has(uri);
    if (closing) {
      depth -= 1;
      if (bodyDepth !== null && depth < bodyDepth) bodyDepth = null;
      continue;
    }
    if (word) {
      const range = RANGE_REASONS[local];
      if (range !== undefined) ranges.add(range);
      if (local === "body") bodyDepth = depth + 1;
      else if (bodyDepth === depth && !BODY_BLOCKS.has(local))
        otherBodyBlocks.add(local);
      if (local === "p") {
        paragraphs += 1;
        const id = PARA_ID.exec(attributes)?.[1];
        if (id === undefined || !/^[0-9a-f]{8}$/iu.test(id) || /^0+$/u.test(id))
          reasons.add("paragraph without a valid paraId");
        else if (paraIds.has(id.toUpperCase()))
          reasons.add("paraId used twice");
        else paraIds.add(id.toUpperCase());
      }
    } else if (bodyDepth === depth) {
      otherBodyBlocks.add(`${prefix ?? ""}${local}`);
    }
    if (!selfClosing) depth += 1;
  }
  if (otherBodyBlocks.size > 0)
    reasons.add("body block other than a paragraph, table or range marker");
  return {
    readable: true,
    paragraphs,
    reasons: [...reasons].toSorted(),
    ranges: [...ranges].toSorted(),
    otherBodyBlocks: [...otherBodyBlocks].toSorted(),
  };
};

const mainPartPath = async (zip: ZipArchive): Promise<string> => {
  const rels = await zip.entry("_rels/.rels")?.text();
  const target = rels
    ?.match(/<Relationship\b[^>]*>/gu)
    // The officeDocument relationship type itself, not the others whose
    // type URIs start with the same namespace (extended properties).
    ?.find((relationship) =>
      /\sType\s*=\s*["'][^"']*\/relationships\/officeDocument["']/u.test(
        relationship,
      ),
    )
    ?.match(/\sTarget\s*=\s*["']([^"']*)["']/u)?.[1];
  return target === undefined
    ? "word/document.xml"
    : target.replace(/^\//u, "");
};

const errorText = (error: unknown) =>
  error instanceof Error ? error.message.slice(0, 80) : "unknown error";

/** Enter at the end of the middle body paragraph, then a selective save. */
const trySplice = async (bytes: Uint8Array): Promise<Trial> => {
  const stamped = (await ensureParaIds(bytes)).docx;
  const session = createCanonicalSession(
    await parseDocx(stamped.slice().buffer, { preloadFonts: false }),
  );
  if (session.isErr()) return { outcome: "not run", reason: "no session" };
  const { value } = session;
  const paragraphs = value.document.package.document.content.flatMap((block) =>
    block.type === "paragraph" && block.paraId !== undefined
      ? [block.paraId]
      : [],
  );
  const blockId = paragraphs.at(Math.floor(paragraphs.length / 2));
  if (blockId === undefined)
    return { outcome: "not run", reason: "no body paragraph" };
  const used = new Set(paragraphs.map((id) => id.toUpperCase()));
  let fresh = 0x1e_57_00_00;
  while (used.has(fresh.toString(16).toUpperCase())) fresh += 1;
  const at = {
    story: "main" as const,
    blockId,
    offset: value.projection.paragraph(blockId)?.text.length ?? 0,
  };
  const state = EditorState.create({ schema, doc: value.projection.doc });
  const commit = value.prepareIntent(state, {
    type: "splitParagraph",
    at,
    to: at,
    newBlockId: fresh.toString(16).toUpperCase(),
  });
  if (commit.isErr()) return { outcome: "not run", reason: "Enter refused" };
  const published = publishCanonicalProjection({
    session: value,
    state,
    commit: commit.value,
  });
  if (published.isErr()) return { outcome: "not run", reason: "Enter refused" };
  const attempts: SelectiveSaveAttempt[] = [];
  const previous = globalThis.__folioLayoutInstrumentation;
  globalThis.__folioLayoutInstrumentation = {
    onSelectiveSaveAttempt: (attempt) => attempts.push(attempt),
  };
  try {
    const saved = await serializeCanonicalSave({
      snapshot: value.captureSaveSnapshot(),
      featureFlags: { selectiveSave: true },
    });
    if (saved.path === "selective") return { outcome: "spliced" };
    const attempt = attempts.at(-1);
    if (attempt === undefined)
      return { outcome: "not run", reason: "selective save not tried" };
    return {
      outcome: "declined",
      reason: attempt.decline?.reason ?? "declined before the splice",
    };
  } finally {
    globalThis.__folioLayoutInstrumentation = previous;
  }
};

const scanFile = async (file: string): Promise<Scan> => {
  let bytes: Uint8Array;
  let scan: Scan;
  try {
    bytes = new Uint8Array(await Bun.file(file).arrayBuffer());
    const zip = await ZipArchive.read(bytes);
    const xml = await zip.entry(await mainPartPath(zip))?.text();
    if (xml === undefined) return { readable: false, reason: "no main part" };
    scan = scanMainPart(xml);
  } catch (error) {
    return { readable: false, reason: errorText(error) };
  }
  if (!scan.readable) return scan;
  try {
    return { ...scan, trial: await trySplice(bytes) };
  } catch (error) {
    return { ...scan, trial: { outcome: "not run", reason: errorText(error) } };
  }
};

const outDirectory = process.argv[2];
if (outDirectory === undefined)
  throw new TypeError("Usage: structural-save-census.ts <out dir>");

const listed = Bun.spawnSync(["git", "ls-files", "*.docx"], {
  cwd: REPOSITORY_ROOT,
});
const checkedIn = listed.stdout
  .toString()
  .split("\n")
  .filter((file) => file !== "")
  .toSorted()
  .map((file) => ({
    set: "checked-in",
    file,
    absolute: path.join(REPOSITORY_ROOT, file),
  }));
const lock = await loadCorpusLock();
const corpus = lock.sources.flatMap((source) =>
  source.files
    .filter(({ path: file }) => file.toLowerCase().endsWith(".docx"))
    .map(({ path: file }) => ({
      set: "public corpus",
      file: `${source.id}/${file}`,
      absolute: path.join(sourceCheckoutPath(source.id), file),
    })),
);
const files = [
  ...checkedIn,
  ...corpus.filter((entry) => existsSync(entry.absolute)),
];

const results: ({ set: string; file: string } & Scan)[] = [];
for (const entry of files) {
  // oxlint-disable-next-line no-await-in-loop -- one document in memory at a time
  results.push({
    set: entry.set,
    file: entry.file,
    ...(await scanFile(entry.absolute)),
  });
}

const lines = [
  "# Structural save census",
  "",
  "Would a structural selective save (after Enter) splice the body, or decline to the full repack? Static reasons from the source body only; a document can have several. Public corpus files not fetched are not counted.",
  "",
];
for (const set of ["checked-in", "public corpus"]) {
  const inSet = results.filter((result) => result.set === set);
  const readable = inSet.flatMap((result) => (result.readable ? [result] : []));
  const declining = readable.filter((result) => result.reasons.length > 0);
  const byReason = new Map<string, number>();
  for (const result of readable)
    for (const reason of result.reasons)
      byReason.set(reason, (byReason.get(reason) ?? 0) + 1);
  const byRange = new Map<string, number>();
  for (const result of readable)
    for (const range of result.ranges)
      byRange.set(range, (byRange.get(range) ?? 0) + 1);
  const onlyParaIds = declining.filter((result) =>
    result.reasons.every((reason) => reason.includes("paraId")),
  ).length;
  const blocks = new Map<string, number>();
  for (const result of readable)
    for (const block of result.otherBodyBlocks)
      blocks.set(block, (blocks.get(block) ?? 0) + 1);
  const percent = (count: number) =>
    readable.length === 0
      ? "n/a"
      : `${((100 * count) / readable.length).toFixed(1)}%`;
  lines.push(
    `## ${set}: ${String(inSet.length)} files, ${String(readable.length)} readable`,
    "",
    `- Would decline: ${String(declining.length)} (${percent(declining.length)}); could splice: ${String(readable.length - declining.length)} (${percent(readable.length - declining.length)}).`,
    `- Declining only for paraIds (a host that stamps ids at ingest avoids these): ${String(onlyParaIds)}.`,
    `- Unreadable: ${String(inSet.length - readable.length)}; readable with no paragraph found by the scan: ${String(readable.filter((result) => result.paragraphs === 0).length)}.`,
    "",
    "| Reason | Files | Share of readable |",
    "|---|---:|---:|",
    ...[...byReason.entries()]
      .toSorted(([, left], [, right]) => right - left)
      .map(
        ([reason, count]) =>
          `| ${reason} | ${String(count)} | ${percent(count)} |`,
      ),
    "",
    "| Holds (carried unless the edit splits one) | Files | Share of readable |",
    "|---|---:|---:|",
    ...[...byRange.entries()]
      .toSorted(([, left], [, right]) => right - left)
      .map(
        ([range, count]) =>
          `| ${range} | ${String(count)} | ${percent(count)} |`,
      ),
    "",
    `Other body blocks: ${
      [...blocks.entries()]
        .toSorted(([, left], [, right]) => right - left)
        .map(([block, count]) => `${block} (${String(count)})`)
        .join(", ") || "none"
    }.`,
    "",
  );
  const trials = readable.flatMap((result) =>
    result.trial === undefined ? [] : [result.trial],
  );
  const ran = trials.filter((trial) => trial.outcome !== "not run");
  const spliced = ran.filter((trial) => trial.outcome === "spliced").length;
  const tally = (outcome: Trial["outcome"]) => {
    const counts = new Map<string, number>();
    for (const trial of trials)
      if (trial.outcome === outcome && "reason" in trial)
        counts.set(trial.reason, (counts.get(trial.reason) ?? 0) + 1);
    return [...counts.entries()].toSorted(
      ([, left], [, right]) => right - left,
    );
  };
  const ofRan = (count: number) =>
    ran.length === 0 ? "n/a" : `${((100 * count) / ran.length).toFixed(1)}%`;
  lines.push(
    `### ${set}: the real save after one Enter (paragraph ids stamped first)`,
    "",
    `- Saved: ${String(ran.length)}; spliced: ${String(spliced)} (${ofRan(spliced)}); declined to the full repack: ${String(ran.length - spliced)} (${ofRan(ran.length - spliced)}); not run: ${String(trials.length - ran.length)}.`,
    "",
    "| Declined because | Files | Share of saved |",
    "|---|---:|---:|",
    ...tally("declined").map(
      ([reason, count]) => `| ${reason} | ${String(count)} | ${ofRan(count)} |`,
    ),
    "",
    `Not run: ${
      tally("not run")
        .slice(0, 10)
        .map(([reason, count]) => `${reason} (${String(count)})`)
        .join(", ") || "none"
    }.`,
    "",
  );
}

mkdirSync(outDirectory, { recursive: true });
writeFileSync(
  path.join(outDirectory, "census.json"),
  `${JSON.stringify(results, null, 2)}\n`,
);
writeFileSync(path.join(outDirectory, "census.md"), `${lines.join("\n")}\n`);
