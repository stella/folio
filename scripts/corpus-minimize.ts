/**
 * Shrink one failing corpus file to the construct that causes the failure.
 *
 * Delta debugging in two passes: first over the package's parts, then over the
 * XML elements of the parts that survived. A candidate counts only if it still
 * fails with the same signature, so the result is a package that reproduces the
 * defect and almost nothing else.
 *
 * The output is a reading aid, not a fixture. It is still third-party content,
 * so it lands in the corpus cache and never in the repository: what belongs in
 * the repository is the synthetic `fast-check` seed a human or an agent writes
 * after reading which construct survived.
 *
 * Usage:
 *   bun scripts/corpus-minimize.ts <source-id>/<relative path> [--invariant <name>]
 *     [--out <dir>] [--budget <evaluations>]
 */

import { mkdir } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";
import JSZip from "jszip";

import { type RunCorpusChecksOptions, runCorpusChecks } from "./lib/corpus-check";
import { deltaDebug } from "./lib/corpus-delta-debug";
import {
  assertOutsideRepository,
  corpusCacheRoot,
  sha256Bytes,
  writeJsonFile,
} from "./lib/corpus-manifest";
import { EXTENDED_CORPUS_INVARIANTS } from "./lib/corpus-invariants/contract";
import {
  CORPUS_INVARIANTS,
  type CorpusFailure,
  type CorpusInvariant,
  failureSignature,
} from "./lib/corpus-signature";
import { prepareXmlPruning } from "./lib/corpus-xml-prune";

class CorpusMinimizeError extends TaggedError("CorpusMinimizeError")<{ message: string }> {}

const DEFAULT_BUDGET = 400;
/** The parts whose elements are worth shrinking; the rest survive or die whole. */
const PRUNABLE_PARTS: readonly string[] = [
  "word/document.xml",
  "word/styles.xml",
  "word/numbering.xml",
  "word/settings.xml",
];

type Package = { names: string[]; contents: Map<string, Uint8Array> };

const readPackage = async (bytes: Uint8Array): Promise<Package> => {
  const archive = await JSZip.loadAsync(bytes);
  const names = Object.keys(archive.files).filter((name) => !archive.files[name]?.dir);
  names.sort();
  const contents = new Map<string, Uint8Array>();
  await Promise.all(
    names.map(async (name) => {
      contents.set(name, await (archive.file(name) as JSZip.JSZipObject).async("uint8array"));
    }),
  );
  return { names, contents };
};

const writePackage = async (
  names: readonly string[],
  contents: ReadonlyMap<string, Uint8Array>,
) => {
  const archive = new JSZip();
  for (const name of names) {
    const data = contents.get(name);
    if (data !== undefined) {
      archive.file(name, data);
    }
  }
  return await archive.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

const failureMatching = (failures: readonly CorpusFailure[], signature: string): boolean =>
  failures.some((failure) => failureSignature(failure) === signature);

/**
 * The minimiser runs one file hundreds of times, so it must never skip an
 * invariant for slowness: a budget that cut a stage short would make the search
 * stop reproducing the signature it is shrinking towards.
 */
const MINIMIZE_BUDGETS = {
  invariantBudgetMs: Number.MAX_SAFE_INTEGER,
  fileBudgetMs: Number.MAX_SAFE_INTEGER,
};

const KNOWN_INVARIANTS = new Set<string>([
  ...Object.values(CORPUS_INVARIANTS),
  ...Object.values(EXTENDED_CORPUS_INVARIANTS),
]);

const asInvariant = (value: string | undefined): CorpusInvariant | undefined => {
  if (value === undefined) {
    return undefined;
  }
  if (!KNOWN_INVARIANTS.has(value)) {
    throw new CorpusMinimizeError({
      message: `--invariant expects one of: ${[...KNOWN_INVARIANTS].sort().join(", ")}`,
    });
  }
  // SAFETY: the value was just proven a member of the invariant union's key set.
  return value as CorpusInvariant;
};

/**
 * Shrinking costs one evaluation per candidate, hundreds of them, and the search
 * only ever asks about one signature. Running the other invariants on every
 * candidate would make shrinking a finding cost more than finding it.
 */
const checkOptions = (invariant: CorpusInvariant | undefined): RunCorpusChecksOptions =>
  invariant === undefined ? MINIMIZE_BUDGETS : { ...MINIMIZE_BUDGETS, only: new Set([invariant]) };

const reproducesSignature = async (
  bytes: Uint8Array,
  signature: string,
  invariant: CorpusInvariant,
): Promise<boolean> => {
  const result = await runCorpusChecks(bytes, checkOptions(invariant));
  return result.kind === "checked" && failureMatching(result.failures, signature);
};

const flagValue = (args: readonly string[], flag: string): string | undefined => {
  const index = args.indexOf(flag);
  if (index === -1) {
    return undefined;
  }
  const value = args.at(index + 1);
  if (value === undefined || value.startsWith("--")) {
    throw new CorpusMinimizeError({ message: `${flag} needs a value` });
  }
  return value;
};

const resolveInput = (target: string): string => {
  if (path.isAbsolute(target)) {
    return target;
  }
  return path.join(corpusCacheRoot(), "sources", target);
};

type MinimizedPart = {
  part: string;
  elementsKept: number;
  elementsTotal: number;
  prunable: boolean;
};

type MinimizeElementsOptions = {
  contents: Map<string, Uint8Array>;
  names: readonly string[];
  signature: string;
  invariant: CorpusInvariant;
  remainingBudget: number;
};

const minimizeElements = async ({
  contents,
  names,
  signature,
  invariant,
  remainingBudget,
}: MinimizeElementsOptions): Promise<{ parts: MinimizedPart[]; evaluations: number }> => {
  const parts: MinimizedPart[] = [];
  let evaluations = 0;
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();

  for (const part of PRUNABLE_PARTS) {
    const original = contents.get(part);
    if (original === undefined || !names.includes(part) || evaluations >= remainingBudget) {
      continue;
    }
    const prunable = prepareXmlPruning(decoder.decode(original));
    const withPart = async (keep: ReadonlySet<number>): Promise<Uint8Array> => {
      contents.set(part, encoder.encode(prunable.render(keep)));
      return await writePackage(names, contents);
    };

    // A rebuild with nothing removed must still fail the same way, or the
    // parser round trip itself changed the input and no deletion below it can
    // be trusted.
    evaluations += 1;
    const probe = await withPart(new Set(prunable.addresses));
    if (!(await reproducesSignature(probe, signature, invariant))) {
      contents.set(part, original);
      parts.push({
        part,
        elementsKept: prunable.addresses.length,
        elementsTotal: prunable.addresses.length,
        prunable: false,
      });
      continue;
    }

    const minimized = await deltaDebug({
      items: prunable.addresses,
      budget: remainingBudget - evaluations,
      reproduces: async (kept) =>
        await reproducesSignature(await withPart(new Set(kept)), signature, invariant),
    });
    evaluations += minimized.evaluations;
    contents.set(part, encoder.encode(prunable.render(new Set(minimized.kept))));
    parts.push({
      part,
      elementsKept: minimized.kept.length,
      elementsTotal: prunable.addresses.length,
      prunable: true,
    });
  }

  return { parts, evaluations };
};

const main = async (args: string[]): Promise<void> => {
  const target = args.at(0);
  if (target === undefined || target.startsWith("--")) {
    throw new CorpusMinimizeError({
      message:
        "Usage: bun scripts/corpus-minimize.ts <source-id>/<path> [--invariant <name>] [--out <dir>] [--budget <n>]",
    });
  }
  const budget = Number(flagValue(args, "--budget") ?? DEFAULT_BUDGET);
  if (!Number.isInteger(budget) || budget < 1) {
    throw new CorpusMinimizeError({ message: "--budget expects a positive integer" });
  }
  const wantedInvariant = flagValue(args, "--invariant");

  const inputPath = resolveInput(target);
  const inputBytes = new Uint8Array(await Bun.file(inputPath).arrayBuffer());
  const baseline = await runCorpusChecks(inputBytes, checkOptions(asInvariant(wantedInvariant)));
  if (baseline.kind !== "checked") {
    throw new CorpusMinimizeError({
      message: `${target} is classified ${baseline.reason}: ${baseline.detail}`,
    });
  }
  const failure = baseline.failures.find(
    (candidate) => wantedInvariant === undefined || candidate.invariant === wantedInvariant,
  );
  if (failure === undefined) {
    throw new CorpusMinimizeError({
      message:
        wantedInvariant === undefined
          ? `${target} passes every invariant; there is nothing to minimise`
          : `${target} does not fail \`${wantedInvariant}\``,
    });
  }
  const signature = failureSignature(failure);
  process.stdout.write(`Reproducing: ${signature}\n`);

  const { names, contents } = await readPackage(inputBytes);
  const partResult = await deltaDebug({
    items: names,
    budget,
    reproduces: async (kept) =>
      await reproducesSignature(await writePackage(kept, contents), signature, failure.invariant),
  });
  process.stdout.write(
    `Parts: ${partResult.kept.length}/${names.length} kept after ${partResult.evaluations} evaluations\n`,
  );

  const elementResult = await minimizeElements({
    contents,
    names: partResult.kept,
    invariant: failure.invariant,
    signature,
    remainingBudget: budget - partResult.evaluations,
  });
  for (const part of elementResult.parts) {
    process.stdout.write(
      part.prunable
        ? `  ${part.part}: ${part.elementsKept}/${part.elementsTotal} elements kept\n`
        : `  ${part.part}: not shrunk; an XML round trip alone changes the failure\n`,
    );
  }

  const minimized = await writePackage(partResult.kept, contents);
  const outDirectory = assertOutsideRepository(
    flagValue(args, "--out") ??
      path.join(corpusCacheRoot(), "minimized", sha256Bytes(inputBytes).slice(0, 12)),
  );
  await mkdir(outDirectory, { recursive: true });
  await Bun.write(path.join(outDirectory, "minimized.docx"), minimized);
  const documentXml = contents.get("word/document.xml");
  if (documentXml !== undefined) {
    await Bun.write(path.join(outDirectory, "document.xml"), documentXml);
  }
  await writeJsonFile(path.join(outDirectory, "summary.json"), {
    source: target,
    sourceSha256: sha256Bytes(inputBytes),
    signature,
    inputBytes: inputBytes.byteLength,
    minimizedBytes: minimized.byteLength,
    partsKept: partResult.kept,
    partsDropped: names.filter((name) => !partResult.kept.includes(name)),
    parts: elementResult.parts,
    evaluations: partResult.evaluations + elementResult.evaluations,
    exhaustedBudget: partResult.exhaustedBudget,
  });
  process.stdout.write(
    `Minimised ${inputBytes.byteLength} -> ${minimized.byteLength} bytes in ${outDirectory}\n`,
  );
};

if (import.meta.main) {
  main(process.argv.slice(2)).catch((cause: unknown) => {
    process.stderr.write(`${cause instanceof Error ? cause.message : String(cause)}\n`);
    process.exitCode = 1;
  });
}
