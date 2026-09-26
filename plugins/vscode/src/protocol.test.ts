import { describe, expect, test } from "bun:test";

import { isHostMessage, isWebviewMessage, pageCountLabel } from "./protocol";

describe("isHostMessage", () => {
  test("accepts each message the extension sends", () => {
    expect(isHostMessage({ type: "loading", fileName: "a.docx" })).toBe(true);
    expect(
      isHostMessage({ type: "document", fileName: "a.docx", html: "<p></p>", pageCount: 3 }),
    ).toBe(true);
    expect(isHostMessage({ type: "error", fileName: "a.docx", message: "bad" })).toBe(true);
    expect(isHostMessage({ type: "error", fileName: "a.docx", message: "bad", hint: "fix" })).toBe(
      true,
    );
  });

  test("refuses a message of the wrong shape", () => {
    expect(isHostMessage(null)).toBe(false);
    expect(isHostMessage("document")).toBe(false);
    expect(isHostMessage({ type: "loading" })).toBe(false);
    expect(isHostMessage({ type: "unknown", fileName: "a.docx" })).toBe(false);
    expect(isHostMessage({ type: "document", fileName: "a.docx", html: "", pageCount: 1.5 })).toBe(
      false,
    );
    expect(isHostMessage({ type: "document", fileName: "a.docx", html: "", pageCount: -1 })).toBe(
      false,
    );
    expect(isHostMessage({ type: "document", fileName: "a.docx", pageCount: 1 })).toBe(false);
    expect(isHostMessage({ type: "error", fileName: "a.docx", message: "bad", hint: 1 })).toBe(
      false,
    );
  });
});

describe("isWebviewMessage", () => {
  test("accepts ready and retry only", () => {
    expect(isWebviewMessage({ type: "ready" })).toBe(true);
    expect(isWebviewMessage({ type: "retry" })).toBe(true);
    expect(isWebviewMessage({ type: "render" })).toBe(false);
    expect(isWebviewMessage([])).toBe(false);
    expect(isWebviewMessage(undefined)).toBe(false);
  });
});

describe("pageCountLabel", () => {
  test("agrees in number", () => {
    expect(pageCountLabel(1)).toBe("1 page");
    expect(pageCountLabel(0)).toBe("0 pages");
    expect(pageCountLabel(12)).toBe("12 pages");
  });
});
