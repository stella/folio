import { describe, expect, test } from "bun:test";

import { applySanitizedImageSrc, sanitizeImageSrc } from "./sanitizeImageSrc";

describe("sanitizeImageSrc", () => {
  test("allows data:image and blob URLs", () => {
    expect(sanitizeImageSrc("data:image/png;base64,abc")).toBe("data:image/png;base64,abc");
    expect(sanitizeImageSrc("data:image/svg+xml;charset=utf-8,%3Csvg/%3E")).toContain(
      "data:image/svg+xml",
    );
    expect(sanitizeImageSrc("blob:https://example.com/uuid")).toBe("blob:https://example.com/uuid");
  });

  test("rejects remote and executable schemes", () => {
    expect(sanitizeImageSrc("https://evil.example/a.png")).toBeUndefined();
    expect(sanitizeImageSrc("http://evil.example/a.png")).toBeUndefined();
    expect(sanitizeImageSrc("javascript:alert(1)")).toBeUndefined();
    expect(sanitizeImageSrc("data:text/html,<script>alert(1)</script>")).toBeUndefined();
    expect(sanitizeImageSrc("file:///etc/passwd")).toBeUndefined();
  });

  test("rejects empty and non-string input", () => {
    expect(sanitizeImageSrc("")).toBeUndefined();
    expect(sanitizeImageSrc("   ")).toBeUndefined();
    expect(sanitizeImageSrc(null)).toBeUndefined();
    expect(sanitizeImageSrc(undefined)).toBeUndefined();
  });
});

describe("applySanitizedImageSrc", () => {
  test("assigns accepted sources", () => {
    const img = { src: "" };
    applySanitizedImageSrc(img, "blob:https://example.com/uuid");
    expect(img.src).toBe("blob:https://example.com/uuid");
  });

  test("leaves src untouched for rejected sources", () => {
    const img = { src: "" };
    applySanitizedImageSrc(img, "https://evil.example/a.png");
    expect(img.src).toBe("");
    applySanitizedImageSrc(img, undefined);
    expect(img.src).toBe("");
  });
});
