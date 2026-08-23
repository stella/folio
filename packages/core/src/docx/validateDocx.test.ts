import { validateDocxPackage } from "@stll/docx-core";
import { describe, expect, test } from "bun:test";

import { createEmptyDocx, validateDocx } from "./rezip";

describe("validateDocx", () => {
  test.each([
    ["generated DOCX", () => createEmptyDocx()],
    ["invalid archive", () => Promise.resolve(new Uint8Array([1, 2, 3]).buffer)],
  ])("uses the canonical package decision for %s", async (_name, createBytes) => {
    const bytes = await createBytes();
    const [canonical, compatibility] = await Promise.all([
      validateDocxPackage(bytes),
      validateDocx(bytes),
    ]);

    expect(compatibility.valid).toBe(canonical.valid);
    expect(compatibility.errors).toEqual(canonical.valid ? [] : [canonical.error]);
    expect(compatibility.warnings).toEqual([]);
  });
});
