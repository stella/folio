#!/usr/bin/env bun
/**
 * Capture Microsoft Word line endpoints into a portable manifest, or validate
 * Folio against a previously captured manifest without launching Word.
 */
import { mkdir } from "node:fs/promises";
import path from "node:path";

import { createFolioExtractor } from "./folioExtract";
import {
  LineEndpointManifestError,
  compareLineEndpoints,
  createWordLineEndpointManifest,
  readLineEndpointManifest,
} from "./lineEndpoints";
import { sha256OfFile } from "./pdfReference";
import type { LineEndpointDivergence } from "./lineEndpoints";
import { getWordTruth } from "./wordTruth";

const EXIT_OK = 0;
const EXIT_DIVERGENT = 1;
const EXIT_INFRA_FAILURE = 2;

export type LineEndpointCliCommand =
  | { type: "help" }
  | { type: "capture"; docxPath: string; outputPath: string; refreshWord: boolean }
  | {
      type: "validate";
      docxPath: string;
      manifestPath: string;
      headed: boolean;
      reuseServer: boolean;
    };

export const parseLineEndpointCliArgs = (argv: string[]): LineEndpointCliCommand => {
  const command = argv.at(0);
  if (command === undefined || command === "--help" || command === "-h") {
    return { type: "help" };
  }
  if (command !== "capture" && command !== "validate") {
    throw new LineEndpointManifestError(`Unknown command: ${command}`);
  }

  const docxPath = argv.at(1);
  if (docxPath === undefined || docxPath.startsWith("--")) {
    throw new LineEndpointManifestError(`${command} requires a DOCX path.`);
  }

  if (command === "capture") {
    let outputPath: string | undefined;
    let refreshWord = false;
    for (let index = 2; index < argv.length; index += 1) {
      const arg = argv[index];
      if (arg === "--refresh-word") {
        refreshWord = true;
        continue;
      }
      if (arg === "--output") {
        outputPath = requireFlagValue({ argv, index, flag: arg });
        index += 1;
        continue;
      }
      throw new LineEndpointManifestError(`Unknown capture option: ${String(arg)}`);
    }
    if (outputPath === undefined) {
      throw new LineEndpointManifestError("capture requires --output <manifest.json>.");
    }
    return { type: "capture", docxPath, outputPath, refreshWord };
  }

  let manifestPath: string | undefined;
  let headed = false;
  let reuseServer = false;
  for (let index = 2; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--headed") {
      headed = true;
      continue;
    }
    if (arg === "--reuse-server") {
      reuseServer = true;
      continue;
    }
    if (arg === "--manifest") {
      manifestPath = requireFlagValue({ argv, index, flag: arg });
      index += 1;
      continue;
    }
    throw new LineEndpointManifestError(`Unknown validate option: ${String(arg)}`);
  }
  if (manifestPath === undefined) {
    throw new LineEndpointManifestError("validate requires --manifest <manifest.json>.");
  }
  return { type: "validate", docxPath, manifestPath, headed, reuseServer };
};

type RequireFlagValueOptions = {
  argv: string[];
  index: number;
  flag: string;
};

const requireFlagValue = ({ argv, index, flag }: RequireFlagValueOptions): string => {
  const value = argv[index + 1];
  if (value === undefined || value.startsWith("--")) {
    throw new LineEndpointManifestError(`${flag} requires a path.`);
  }
  return value;
};

const HELP_TEXT = `Microsoft Word line-endpoint validator

Usage:
  bun parity/lineEndpointsCli.ts capture <file.docx> --output <manifest.json> [--refresh-word]
  bun parity/lineEndpointsCli.ts validate <file.docx> --manifest <manifest.json> [--headed] [--reuse-server]

Capture requires Microsoft Word for Mac and mutool. Validate uses only the
captured manifest and Folio, so it can run without Word. Manifests contain
normalized document line text; do not commit captures from confidential files.
`;

const assertFileExists = async (filePath: string, label: string): Promise<void> => {
  if (!(await Bun.file(filePath).exists())) {
    throw new LineEndpointManifestError(`${label} does not exist: ${filePath}`);
  }
};

const capture = async (
  command: Extract<LineEndpointCliCommand, { type: "capture" }>,
): Promise<number> => {
  const docxPath = path.resolve(command.docxPath);
  const outputPath = path.resolve(command.outputPath);
  await assertFileExists(docxPath, "DOCX");

  const [reference, sourceSha256] = await Promise.all([
    getWordTruth(docxPath, { refresh: command.refreshWord }),
    sha256OfFile(docxPath),
  ]);
  const manifest = createWordLineEndpointManifest({
    reference,
    sourceFileName: path.basename(docxPath),
    sourceSha256,
  });

  await mkdir(path.dirname(outputPath), { recursive: true });
  await Bun.write(outputPath, `${JSON.stringify(manifest, null, 2)}\n`);
  const lineCount = manifest.pages.reduce((total, page) => total + page.lines.length, 0);
  console.log(
    `Captured ${lineCount} Word line endpoint${lineCount === 1 ? "" : "s"} across ${manifest.pages.length} page${manifest.pages.length === 1 ? "" : "s"}: ${outputPath}`,
  );
  return EXIT_OK;
};

const validate = async (
  command: Extract<LineEndpointCliCommand, { type: "validate" }>,
): Promise<number> => {
  const docxPath = path.resolve(command.docxPath);
  const manifestPath = path.resolve(command.manifestPath);
  await Promise.all([
    assertFileExists(docxPath, "DOCX"),
    assertFileExists(manifestPath, "Manifest"),
  ]);

  const [manifest, sourceSha256] = await Promise.all([
    readLineEndpointManifest(manifestPath),
    sha256OfFile(docxPath),
  ]);
  if (manifest.source.sha256 !== sourceSha256) {
    throw new LineEndpointManifestError(
      `Manifest source hash does not match ${path.basename(docxPath)}. Capture a new manifest for this exact DOCX.`,
    );
  }

  const extractor = await createFolioExtractor({
    headless: !command.headed,
    reuseServer: command.reuseServer,
  });
  try {
    const folio = await extractor.extract(docxPath);
    const result = compareLineEndpoints(manifest, folio.geom);
    if (result.matches) {
      console.log(
        `PASS: ${result.folioLines} Folio line endpoints match Word ${manifest.reference.wordVersion ?? "unknown version"}.`,
      );
      return EXIT_OK;
    }

    console.error(
      `FAIL: ${result.divergences.length} line-endpoint divergence${result.divergences.length === 1 ? "" : "s"} (${result.referenceLines} Word lines, ${result.folioLines} Folio lines).`,
    );
    for (const divergence of result.divergences) {
      console.error(`  ${formatDivergence(divergence)}`);
    }
    return EXIT_DIVERGENT;
  } finally {
    await extractor.close();
  }
};

const formatDivergence = (divergence: LineEndpointDivergence): string => {
  switch (divergence.kind) {
    case "page-count":
      return `page count: Word ${divergence.reference}, Folio ${divergence.folio}`;
    case "pagination":
      return `pagination: “${divergence.text}” on Word page ${divergence.referencePage}, Folio page ${divergence.folioPage}`;
    case "line-break":
      return `page ${divergence.page}: Word [${divergence.referenceTexts.join(" | ")}], Folio [${divergence.folioTexts.join(" | ")}]`;
    case "missing-line":
      return `page ${divergence.page}: missing Folio line “${divergence.text}”`;
    case "extra-line":
      return `page ${divergence.page}: extra Folio line “${divergence.text}”`;
    case "text-mismatch":
      return `page ${divergence.page}: Word “${divergence.referenceText}”, Folio “${divergence.folioText}”`;
  }
};

export const runLineEndpointCli = async (argv: string[]): Promise<number> => {
  const command = parseLineEndpointCliArgs(argv);
  if (command.type === "help") {
    console.log(HELP_TEXT);
    return EXIT_OK;
  }
  if (command.type === "capture") return await capture(command);
  return await validate(command);
};

if (import.meta.main) {
  try {
    process.exitCode = await runLineEndpointCli(process.argv.slice(2));
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    console.error(`Word line-endpoint validator failed: ${err.name}: ${err.message}`);
    process.exitCode = EXIT_INFRA_FAILURE;
  }
}
