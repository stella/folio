/**
 * The image table's contract is that an unusable source becomes a reported
 * failure, never an exception: a builder that throws paints no page at all.
 */

import { describe, expect, test } from "bun:test";

import { ImageTable } from "./imagePrimitives";

const MALFORMED_BASE64_URL = "data:image/png;base64,!!!not-base64!!!";

describe("ImageTable.intern", () => {
  test("a malformed base64 payload is reported, not thrown", () => {
    // Pins that the payload really is the hazardous class: `atob` rejects it,
    // so this exercises the decode failure rather than an empty result.
    expect(() => atob("!!!not-base64!!!")).toThrow();

    const images = new ImageTable();
    expect(images.intern(MALFORMED_BASE64_URL)).toBeUndefined();
    expect(images.failureFor(MALFORMED_BASE64_URL)).toContain("base64");
    expect(images.snapshot()).toHaveLength(0);
  });

  test("the failure is cached, so a repeated source stays a reported failure", () => {
    const images = new ImageTable();
    images.intern(MALFORMED_BASE64_URL);
    expect(images.intern(MALFORMED_BASE64_URL)).toBeUndefined();
  });

  test("a well-formed PNG data: URL interns", () => {
    // 1x1 transparent PNG.
    const png =
      "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
    const images = new ImageTable();
    expect(images.intern(png)).toBe(0);
    expect(images.snapshot().at(0)?.format).toBe("png");
  });
});
