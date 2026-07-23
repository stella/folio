/**
 * Benchmark corpus loader.
 *
 * Loads a representative handful of real `.docx` fixtures already checked into
 * the repo and hands them to the bench files as in-memory bytes. Everything
 * here is runtime-agnostic on purpose: fixtures are read with `node:fs` /
 * `node:path` (never `Bun.file`) so the same bench files run under Bun locally
 * (`bun run bench`) and under Node when CodSpeed instruments them in CI.
 *
 * We pick three documents spanning the size range rather than running all ~30
 * corpus files per case, so the suite stays fast and the numbers stay stable:
 *
 *   - small  ~13 KB  the docx-editor demo document
 *   - medium ~39 KB  a mixed-content sample
 *   - large  ~180 KB a long, image-heavy real-world document
 *
 * The other fixtures under
 * `packages/core/src/docx/__tests__/__fixtures__/{corpus,regressions}` remain
 * available if a future bench wants to target a specific edge case.
 */

import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, "..");

export type FixtureSize = "small" | "medium" | "large";

export type Fixture = {
  /** Stable label used to group bench cases (also the size bucket). */
  readonly size: FixtureSize;
  /** Human-readable name (the file's basename). */
  readonly name: string;
  /** Byte length of the document. */
  readonly bytes: number;
  /** Cached file contents; never hand this to a parser directly — copy first. */
  readonly buffer: Buffer;
};

type FixtureSpec = {
  size: FixtureSize;
  /** Path relative to the repo root. */
  path: string;
};

const SPECS: readonly FixtureSpec[] = [
  { size: "small", path: "tests/visual/fixtures/docx-editor-demo.docx" },
  { size: "medium", path: "tests/visual/fixtures/sample.docx" },
  { size: "large", path: "tests/visual/fixtures/podily-bps.docx" },
];

function loadFixture(spec: FixtureSpec): Fixture {
  const absolute = resolve(REPO_ROOT, spec.path);
  const buffer = readFileSync(absolute);
  const name = spec.path.slice(spec.path.lastIndexOf("/") + 1);
  return { size: spec.size, name, bytes: buffer.byteLength, buffer };
}

/** The benchmark corpus, ordered small → large. */
export const FIXTURES: readonly Fixture[] = SPECS.map(loadFixture);

/**
 * A fresh `ArrayBuffer` copy of the fixture bytes. Each parse iteration gets
 * its own copy so a parser that retains or transfers its input cannot affect
 * the next iteration.
 */
export function freshArrayBuffer(fixture: Fixture): ArrayBuffer {
  const { buffer } = fixture;
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

/** Short "label (NN KB)" string for bench case names. */
export function fixtureLabel(fixture: Fixture): string {
  return `${fixture.size} · ${fixture.name} (${Math.round(fixture.bytes / 1024)} KB)`;
}
