import { describe, expect, test } from "bun:test";

import sourceManifest from "../specifications/sources.json";
import { validateSpecificationSourceManifest } from "./specification-sources";

const cloneManifest = (): unknown => structuredClone(sourceManifest);

describe("specification source manifest", () => {
  test("the committed manifest is valid", () => {
    expect(validateSpecificationSourceManifest(sourceManifest)).toEqual([]);
  });

  test("rejects duplicate source ids", () => {
    const manifest = cloneManifest();
    if (typeof manifest !== "object" || manifest === null || !("sources" in manifest)) {
      throw new TypeError("invalid test fixture");
    }
    if (!Array.isArray(manifest.sources)) {
      throw new TypeError("invalid test fixture sources");
    }
    manifest.sources.push(structuredClone(manifest.sources.at(0)));
    expect(validateSpecificationSourceManifest(manifest)).toContain(
      "manifest.sources: duplicate source ids",
    );
  });

  test("rejects source order drift", () => {
    const manifest = cloneManifest();
    if (typeof manifest !== "object" || manifest === null || !("sources" in manifest)) {
      throw new TypeError("invalid test fixture");
    }
    if (!Array.isArray(manifest.sources)) {
      throw new TypeError("invalid test fixture sources");
    }
    manifest.sources.reverse();
    expect(validateSpecificationSourceManifest(manifest)).toContain(
      "manifest.sources: sources must be sorted by id",
    );
  });

  test("rejects an invalid artifact digest", () => {
    const manifest = cloneManifest();
    if (typeof manifest !== "object" || manifest === null || !("sources" in manifest)) {
      throw new TypeError("invalid test fixture");
    }
    if (!Array.isArray(manifest.sources)) {
      throw new TypeError("invalid test fixture sources");
    }
    const source = manifest.sources.at(0);
    if (typeof source !== "object" || source === null) {
      throw new TypeError("invalid test fixture source");
    }
    Object.assign(source, { sha256: "invalid" });
    expect(validateSpecificationSourceManifest(manifest)).toContain(
      "sources[0].sha256: expected a lowercase SHA-256 digest",
    );
  });

  test("rejects cache paths outside the specification cache", () => {
    const manifest = cloneManifest();
    if (typeof manifest !== "object" || manifest === null || !("sources" in manifest)) {
      throw new TypeError("invalid test fixture");
    }
    if (!Array.isArray(manifest.sources)) {
      throw new TypeError("invalid test fixture sources");
    }
    const source = manifest.sources.at(0);
    if (typeof source !== "object" || source === null) {
      throw new TypeError("invalid test fixture source");
    }
    Object.assign(source, { cachePath: ".cache/specifications/../outside" });
    expect(validateSpecificationSourceManifest(manifest)).toContain(
      "sources[0].cachePath: expected a path under .cache/specifications",
    );
  });
});
