/**
 * The repository's DOCX fixtures, read once, for golden comparisons of every
 * XML reader in the workspace.
 *
 * A golden records what a reader answered for each fixture part when the
 * answer was recorded; a test recomputes the answer and compares it. Answers
 * are stored as SHA-256 digests of their JSON so a golden for a whole package
 * stays a few lines long while still failing on any difference.
 */
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

export const REPO_ROOT = resolve(import.meta.dir, "../..");

export const DOCX_FIXTURE_GLOBS = [
  "tests/visual/fixtures/*.docx",
  "packages/core/src/docx/__fixtures__/*.docx",
  "packages/core/src/docx/__tests__/__fixtures__/**/*.docx",
] as const;

const XML_PART = /\.(?:xml|rels)$/iu;

export type FixturePackage = {
  /** Repository-relative path of the `.docx`. */
  path: string;
  bytes: Uint8Array<ArrayBuffer>;
  /** Every `.xml` and `.rels` part, decoded as UTF-8, keyed by part path. */
  parts: ReadonlyMap<string, string>;
  /** Every entry path in the archive. */
  paths: ReadonlySet<string>;
};

export const readFixturePackages = async (): Promise<FixturePackage[]> => {
  const files = DOCX_FIXTURE_GLOBS.flatMap((pattern) => [
    ...new Bun.Glob(pattern).scanSync({ cwd: REPO_ROOT }),
  ]).toSorted();
  const packages: FixturePackage[] = [];
  for (const path of files) {
    const bytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, path)));
    const zip = await ZipArchive.read(bytes);
    const parts = new Map<string, string>();
    for (const entry of zip.entries()) {
      if (XML_PART.test(entry.path)) parts.set(entry.path, await entry.text());
    }
    packages.push({ path, bytes, parts, paths: new Set(zip.paths()) });
  }
  return packages;
};

type DescribedElement = {
  type?: string;
  namespaceUri?: string;
  namespaceScope?: DescribedScope;
  elements?: DescribedElement[];
};
type DescribedScope = {
  bindings: ReadonlyMap<string, string>;
  parent?: DescribedScope;
};

/** Every binding in scope, flattened, so two chains that resolve alike compare equal. */
const scopeKey = (scope: DescribedScope | undefined): string => {
  const resolved = new Map<string, string>();
  const chain: DescribedScope[] = [];
  for (let current = scope; current; current = current.parent) {
    chain.push(current);
  }
  for (const link of chain.toReversed()) {
    for (const [prefix, uri] of link.bindings) resolved.set(prefix, uri);
  }
  return JSON.stringify(
    [...resolved].toSorted(([left], [right]) => left.localeCompare(right)),
  );
};

/**
 * A parsed tree as JSON, including each element's resolved namespace and the
 * bindings in scope, which are non-enumerable and so invisible to `toEqual`.
 */
export const describeXmlTree = (root: DescribedElement): string =>
  JSON.stringify(root, (key, value: DescribedElement) =>
    key !== "" && value?.type === "element"
      ? {
          ...value,
          namespace: value.namespaceUri,
          scope: scopeKey(value.namespaceScope),
        }
      : value,
  );

/** A stable digest of a JSON-serialisable answer. */
export const answerDigest = (answer: unknown): string =>
  createHash("sha256")
    .update(JSON.stringify(answer) ?? "undefined")
    .digest("hex");
