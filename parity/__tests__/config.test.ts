import { describe, expect, test } from "bun:test";
import path from "node:path";

import { CACHE_DIR, resolveReferenceCacheDir } from "../config";

describe("reference cache directory", () => {
  test("defaults to the checkout cache", () => {
    expect(resolveReferenceCacheDir(undefined)).toBe(CACHE_DIR);
    expect(resolveReferenceCacheDir("  ")).toBe(CACHE_DIR);
  });

  test("uses an override as an absolute path", () => {
    expect(resolveReferenceCacheDir("/shared/reference-cache")).toBe("/shared/reference-cache");
    expect(resolveReferenceCacheDir("shared-cache")).toBe(path.resolve("shared-cache"));
  });
});
