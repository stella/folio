import { describe, expect, test } from "bun:test";

import {
  CORPUS_INVARIANTS,
  NO_FRAME,
  describeError,
  failureFromError,
  failureSignature,
  normalizeFailureMessage,
  topFolioFrame,
} from "./lib/corpus-signature";

describe("normalizeFailureMessage", () => {
  test("erases the per-file particulars that split one defect into many", () => {
    const first = normalizeFailureMessage("Duplicate comment id 12 in /Users/a/b/one.docx");
    const second = normalizeFailureMessage("Duplicate comment id 4096 in /home/ci/two.docx");
    expect(first).toBe(second);
  });

  test("keeps a relative part name, which says which part broke", () => {
    expect(normalizeFailureMessage("word/numbering.xml is missing")).toBe(
      "word/numbering.xml is missing",
    );
  });

  test("erases GUIDs and long hex runs", () => {
    expect(
      normalizeFailureMessage("block 0f8e1a2b3c4d5e6f in 6f9619ff-8b86-d011-b42d-00c04fc964ff"),
    ).toBe("block <hex> in <guid>");
  });

  test("collapses newlines so a multi-line error is one signature", () => {
    expect(normalizeFailureMessage("first line\n  second   line\n")).toBe("first line second line");
  });

  test("bounds the length so one verbose error cannot dominate a report", () => {
    const normalized = normalizeFailureMessage("x".repeat(500));
    expect(normalized.length).toBe(161);
    expect(normalized.endsWith("…")).toBe(true);
  });
});

describe("topFolioFrame", () => {
  test("names the innermost frame inside a folio package", () => {
    const stack = [
      "Error: boom",
      "    at panic (/repo/node_modules/better-result/dist/index.mjs:2:322)",
      "    at assertStyleNumberingReferences (/repo/packages/core/src/docx/rezip.ts:3015:11)",
      "    at runCorpusChecks (/repo/scripts/lib/corpus-check.ts:120:5)",
    ].join("\n");
    expect(topFolioFrame(stack)).toBe(
      "packages/core/src/docx/rezip.ts:assertStyleNumberingReferences",
    );
  });

  test("skips a packaged dependency that happens to live under node_modules/packages", () => {
    const stack = "Error: boom\n    at f (/repo/node_modules/x/packages/core/src/a.ts:1:1)";
    expect(topFolioFrame(stack)).toBe(NO_FRAME);
  });

  test("reports no frame when the stack never enters a package", () => {
    expect(topFolioFrame("Error: boom\n    at /repo/scripts/lib/corpus-check.ts:1:1")).toBe(
      NO_FRAME,
    );
    expect(topFolioFrame(undefined)).toBe(NO_FRAME);
  });

  test("handles a frame with no function name", () => {
    expect(topFolioFrame("Error: boom\n    at /repo/packages/core/src/a.ts:3:9")).toBe(
      "packages/core/src/a.ts",
    );
  });
});

describe("describeError", () => {
  test("prefixes the class, because the class is part of the defect", () => {
    class DocxParseError extends Error {
      override name = "DocxParseError";
    }
    expect(describeError(new DocxParseError("bad part"))).toBe("DocxParseError: bad part");
  });

  test("does not double the prefix when the message already carries it", () => {
    const error = new TypeError("TypeError: already prefixed");
    expect(describeError(error)).toBe("TypeError: already prefixed");
  });

  test("stringifies a non-error throw", () => {
    expect(describeError("plain string")).toBe("plain string");
  });
});

describe("failureSignature", () => {
  test("two files failing the same way share a signature", () => {
    const left = failureFromError(
      CORPUS_INVARIANTS.parse,
      Object.assign(new Error("Duplicate comment id 1"), {
        stack: "Error: x\n    at parseDocx (/repo/packages/core/src/docx/parser.ts:10:1)",
      }),
    );
    const right = failureFromError(
      CORPUS_INVARIANTS.parse,
      Object.assign(new Error("Duplicate comment id 9999"), {
        stack: "Error: x\n    at parseDocx (/repo/packages/core/src/docx/parser.ts:12:3)",
      }),
    );
    expect(failureSignature(left)).toBe(failureSignature(right));
  });

  test("the same message under a different invariant is a different signature", () => {
    const error = new Error("same");
    expect(failureSignature(failureFromError(CORPUS_INVARIANTS.parse, error))).not.toBe(
      failureSignature(failureFromError(CORPUS_INVARIANTS.fixedPoint, error)),
    );
  });
});
