/** Stage a bounded tier-one corpus slice as external, hash-addressed flow starts. */
import { appendFile, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { panic } from "better-result";

import { corpusCacheRoot, CORPUS_TIERS, loadCorpusLock, sha256Bytes } from "./lib/corpus-manifest";

import {
  MAX_PUBLIC_CORPUS_BYTES,
  MAX_PUBLIC_CORPUS_FILES,
} from "../test/consumer-scenarios/support/public-corpus";

const lock = await loadCorpusLock();
const cache = corpusCacheRoot();
const directory = path.join(cache, "consumer-flow-starts");
const selected = lock.sources
  .filter(({ tier }) => tier === CORPUS_TIERS.permissive)
  .flatMap(({ id, files }) => files.map((file) => ({ source: id, ...file })))
  .filter(({ bytes }) => bytes >= 1_000 && bytes <= MAX_PUBLIC_CORPUS_BYTES)
  .toSorted((left, right) => left.sha256.localeCompare(right.sha256))
  .filter((file, index, files) => index === 0 || files.at(index - 1)?.sha256 !== file.sha256)
  .slice(0, MAX_PUBLIC_CORPUS_FILES);
if (selected.length !== MAX_PUBLIC_CORPUS_FILES)
  panic("Public corpus flow staging requires eight eligible pinned fixtures");
await mkdir(directory, { recursive: true });
for (const file of selected) {
  const source = path.join(cache, "sources", file.source, file.path);
  const bytes = new Uint8Array(await readFile(source));
  if (bytes.byteLength !== file.bytes || sha256Bytes(bytes) !== file.sha256) {
    panic("Public corpus cache does not match the lock; run bun run corpus:fetch", {
      sha256: file.sha256,
    });
  }
  await writeFile(path.join(directory, `${file.sha256}.docx`), bytes);
}
await writeFile(
  path.join(directory, "index.json"),
  `${JSON.stringify(selected.map(({ sha256 }) => sha256))}\n`,
);
console.log(`Staged ${selected.length} tier-one fixtures outside the checkout: ${directory}`);
if (process.env["GITHUB_ENV"]) {
  await appendFile(process.env["GITHUB_ENV"], `FOLIO_SCENARIO_PUBLIC_CORPUS_DIR=${directory}\n`);
}
