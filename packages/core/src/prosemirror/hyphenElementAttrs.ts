import { panic } from "better-result";
import type { Mark } from "prosemirror-model";

import type { ReadProseMirrorAttrsResult } from "./attrs";

type HyphenElementMarkAttrs = {
  /** Tells this carrier from its neighbours. Monotonic within one paragraph. */
  id: number;
};

export const readHyphenElementMarkAttrs = (
  mark: Mark,
): ReadProseMirrorAttrsResult<HyphenElementMarkAttrs> => {
  const id: unknown = mark.attrs["id"];
  if (mark.type.name !== "hyphenElement") {
    return {
      ok: false,
      issues: [
        {
          path: "mark.type.name",
          message: `Expected hyphenElement, got ${mark.type.name}.`,
        },
      ],
    };
  }
  if (typeof id !== "number" || !Number.isSafeInteger(id) || id < 0) {
    return {
      ok: false,
      issues: [
        {
          path: "hyphenElement.attrs.id",
          message: "Expected a non-negative safe integer.",
        },
      ],
    };
  }
  return { ok: true, value: { id } };
};

export const expectHyphenElementMarkAttrs = (
  mark: Mark,
): HyphenElementMarkAttrs => {
  const result = readHyphenElementMarkAttrs(mark);
  if (!result.ok) {
    const details = result.issues
      .map((issue) => `${issue.path}: ${issue.message}`)
      .join("\n");
    panic(`Invalid ProseMirror hyphen element attrs:\n${details}`);
  }
  return result.value;
};
