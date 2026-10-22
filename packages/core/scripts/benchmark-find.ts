import { panic } from "better-result";
import JSZip from "jszip";
/** Run with Bun: selects the largest fixture by uncompressed document.xml bytes. */
import { relative, resolve } from "node:path";

import { parseDocx } from "../src/docx/parser";
import { toProseDoc } from "../src/prosemirror/conversion/toProseDoc";
import { resolveDocumentFindRanges } from "../src/utils/findMatchView";
import {
  createDefaultFindOptions,
  findInDocument,
} from "../src/utils/findReplace";

const repositoryRoot = resolve(import.meta.dir, "../../..");
const fixtureRoots = [
  "packages/core/src/docx/__tests__/__fixtures__",
  "tests/visual/fixtures",
  "parity/fixtures",
];
const query = process.argv.at(2) ?? "a";
const sampleCount = 15;

type FixtureCandidate = {
  path: string;
  documentXmlBytes: number;
  compressedBytes: number;
};
let largest: FixtureCandidate | null = null;
let fixtureCount = 0;
for (const directory of fixtureRoots) {
  const glob = new Bun.Glob("**/*.docx");
  for await (const path of glob.scan({
    cwd: resolve(repositoryRoot, directory),
    absolute: true,
  })) {
    const bytes = await Bun.file(path).arrayBuffer();
    const zip = await JSZip.loadAsync(bytes);
    const documentXml = zip.file("word/document.xml");
    if (!documentXml)
      panic("Fixture has no document.xml", {
        fixture: relative(repositoryRoot, path),
      });
    const documentXmlBytes = (await documentXml.async("uint8array")).length;
    fixtureCount++;
    if (!largest || documentXmlBytes > largest.documentXmlBytes) {
      largest = { path, documentXmlBytes, compressedBytes: bytes.byteLength };
    }
  }
}
if (!largest) panic("No DOCX fixtures found");

const parseStart = performance.now();
const document = await parseDocx(await Bun.file(largest.path).arrayBuffer(), {
  detectVariables: false,
  preloadFonts: false,
});
const parseMilliseconds = performance.now() - parseStart;
const conversionStart = performance.now();
const bodyDoc = toProseDoc(document);
const conversionMilliseconds = performance.now() - conversionStart;
const options = { ...createDefaultFindOptions(), includeComments: true };
const searchSamples: number[] = [];
const resolutionSamples: number[] = [];
let matchCount = 0;
let rangeCount = 0;
// The first pass warms the code paths; measured samples include fresh matching
// and secondary-story conversions, with the live body PM document reused.
for (let sample = 0; sample <= sampleCount; sample++) {
  const searchStart = performance.now();
  const matches = findInDocument(document, query, options);
  const searchMilliseconds = performance.now() - searchStart;
  const resolutionStart = performance.now();
  const ranges = resolveDocumentFindRanges({ document, bodyDoc, matches });
  const resolutionMilliseconds = performance.now() - resolutionStart;
  if (sample > 0) {
    searchSamples.push(searchMilliseconds);
    resolutionSamples.push(resolutionMilliseconds);
  }
  matchCount = matches.length;
  rangeCount = ranges.length;
}

const summarize = (samples: number[]) => {
  const sorted = samples.toSorted((left, right) => left - right);
  return {
    medianMilliseconds: sorted.at(Math.floor(sorted.length / 2)),
    p95Milliseconds: sorted.at(Math.ceil(sorted.length * 0.95) - 1),
  };
};

console.log(
  JSON.stringify(
    {
      fixture: relative(repositoryRoot, largest.path),
      fixtureCount,
      documentXmlBytes: largest.documentXmlBytes,
      compressedBytes: largest.compressedBytes,
      query,
      sampleCount,
      matchCount,
      rangeCount,
      parseMilliseconds,
      conversionMilliseconds,
      search: summarize(searchSamples),
      resolution: summarize(resolutionSamples),
    },
    null,
    2,
  ),
);
