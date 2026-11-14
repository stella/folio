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
 * a bookmark, field character, permission range or move range anywhere in it,
 * a paragraph without a valid unique paraId, or a body block that is neither
 * a paragraph nor a table. This scans each document's main part for those
 * reasons, mirroring the patch's static checks; the reasons that depend on
 * the edit (a table changed shape, the final section changed) are not here.
 *
 * Usage: bun scripts/structural-save-census.ts <out dir>
 * Scans the checked-in .docx files and, when it has been fetched, the pinned
 * public corpus; writes <out dir>/census.json and <out dir>/census.md.
 */
import { ZipArchive } from "@stll/docx-core/zip";

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
const BODY_BLOCKS = new Set(["p", "tbl", "sectPr"]);
const PARA_ID = /\s[\w.-]+:paraId\s*=\s*["']([^"']*)["']/u;
const TAG = /<(\/?)([\w.-]+:)?([\w.-]+)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/gu;
const NAMESPACE = /\sxmlns(?::([\w.-]+))?\s*=\s*["']([^"']*)["']/gu;

type Scan =
  | { readable: false; reason: string }
  | { readable: true; reasons: string[]; otherBodyBlocks: string[] };

/** The static reasons a structural splice of this main part declines. */
const scanMainPart = (xml: string): Scan => {
  const prefixes = new Map<string, string>();
  for (const [, prefix, uri] of xml.matchAll(NAMESPACE))
    if (!prefixes.has(prefix ?? "")) prefixes.set(prefix ?? "", uri ?? "");
  const reasons = new Set<string>();
  const otherBodyBlocks = new Set<string>();
  const paraIds = new Set<string>();
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
      if (range !== undefined) reasons.add(range);
      if (local === "body") bodyDepth = depth + 1;
      else if (bodyDepth === depth && !BODY_BLOCKS.has(local))
        otherBodyBlocks.add(local);
      if (local === "p") {
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
    reasons.add("body block other than a paragraph or table");
  return {
    readable: true,
    reasons: [...reasons].toSorted(),
    otherBodyBlocks: [...otherBodyBlocks].toSorted(),
  };
};

const mainPartPath = async (zip: ZipArchive): Promise<string> => {
  const rels = await zip.entry("_rels/.rels")?.text();
  const target = rels
    ?.match(/<Relationship\b[^>]*>/gu)
    ?.find((relationship) => relationship.includes("/officeDocument"))
    ?.match(/\sTarget\s*=\s*["']([^"']*)["']/u)?.[1];
  return target === undefined
    ? "word/document.xml"
    : target.replace(/^\//u, "");
};

const scanFile = async (file: string): Promise<Scan> => {
  try {
    const zip = await ZipArchive.read(
      new Uint8Array(await Bun.file(file).arrayBuffer()),
    );
    const xml = await zip.entry(await mainPartPath(zip))?.text();
    if (xml === undefined) return { readable: false, reason: "no main part" };
    return scanMainPart(xml);
  } catch (error) {
    return {
      readable: false,
      reason: error instanceof Error ? error.message.slice(0, 80) : "unread",
    };
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
    `- Unreadable: ${String(inSet.length - readable.length)}.`,
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
    `Other body blocks: ${
      [...blocks.entries()]
        .toSorted(([, left], [, right]) => right - left)
        .map(([block, count]) => `${block} (${String(count)})`)
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
