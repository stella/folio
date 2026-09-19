import { describe, expect, test } from "bun:test";

import corpusManifest from "../corpus/sources.json";
import { validateCorpusManifest } from "./lib/corpus-manifest";

type MutableManifest = { sources: Record<string, unknown>[] };

const cloneManifest = (): MutableManifest => {
  const manifest = structuredClone(corpusManifest) as unknown;
  if (
    typeof manifest !== "object" ||
    manifest === null ||
    !("sources" in manifest) ||
    !Array.isArray(manifest.sources)
  ) {
    throw new TypeError("invalid test fixture");
  }
  return manifest as MutableManifest;
};

const firstSource = (manifest: MutableManifest): Record<string, unknown> => {
  const source = manifest.sources.at(0);
  if (source === undefined) {
    throw new TypeError("invalid test fixture source");
  }
  return source;
};

describe("corpus source manifest", () => {
  test("the committed manifest is valid", () => {
    expect(validateCorpusManifest(corpusManifest)).toEqual([]);
  });

  test("every source records a reviewed licence", () => {
    for (const source of cloneManifest().sources) {
      expect(source["license"]).toMatchObject({
        auditStatus: "reviewed",
        redistribution: "cache-only",
      });
    }
  });

  test("rejects duplicate source ids", () => {
    const manifest = cloneManifest();
    manifest.sources.push(structuredClone(firstSource(manifest)));
    expect(validateCorpusManifest(manifest)).toContain("manifest.sources: duplicate source ids");
  });

  test("rejects source order drift", () => {
    const manifest = cloneManifest();
    manifest.sources.reverse();
    expect(validateCorpusManifest(manifest)).toContain(
      "manifest.sources: sources must be sorted by id",
    );
  });

  test("rejects an unpinned commit", () => {
    const manifest = cloneManifest();
    firstSource(manifest)["commit"] = "main";
    expect(validateCorpusManifest(manifest)).toContain(
      "sources[0].commit: expected a lowercase Git object ID",
    );
  });

  test("rejects a path pattern that would take more than .docx files", () => {
    const manifest = cloneManifest();
    firstSource(manifest)["paths"] = ["/test-data/**"];
    expect(validateCorpusManifest(manifest)).toContain(
      "sources[0].paths[0]: expected a rooted `*.docx` sparse-checkout pattern",
    );
  });

  test("rejects a path pattern that escapes the repository", () => {
    const manifest = cloneManifest();
    firstSource(manifest)["paths"] = ["/../../*.docx"];
    expect(validateCorpusManifest(manifest)).toContain(
      "sources[0].paths[0]: must not traverse upwards",
    );
  });

  test("rejects a redistribution policy other than cache-only", () => {
    const manifest = cloneManifest();
    const license = firstSource(manifest)["license"];
    if (typeof license !== "object" || license === null) {
      throw new TypeError("invalid test fixture licence");
    }
    Object.assign(license, { redistribution: "license-compliant" });
    expect(validateCorpusManifest(manifest)).toContain(
      "sources[0].license.redistribution: corpus content is cache-only",
    );
  });

  test("rejects a non-GitHub repository", () => {
    const manifest = cloneManifest();
    firstSource(manifest)["repository"] = "https://example.com/some/repo";
    expect(validateCorpusManifest(manifest)).toContain(
      "sources[0].repository: expected a GitHub HTTPS repository URL",
    );
  });
});
