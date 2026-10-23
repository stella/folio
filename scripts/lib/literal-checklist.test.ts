import { describe, expect, test } from "bun:test";

import { createLiteralChecklist } from "./literal-checklist";

describe("literal bundle checklists", () => {
  test("reports overlapping rules and sentinels in declared order", () => {
    const patterns = [
      "__SECRET_INTERNALS_DO_NOT_USE",
      "absent",
      "__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED",
      ".folio-root",
      ".folio-root",
    ];
    expect(
      createLiteralChecklist(patterns)(
        ".folio-root __SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED",
      ),
    ).toEqual({
      present: [patterns[0], patterns[2], ".folio-root", ".folio-root"],
      missing: ["absent"],
    });
  });

  test("keeps literals, accents, case, inflections and empty lists exact", () => {
    const scan = createLiteralChecklist([
      ".a+b[0]",
      "Novák",
      "Nováka",
      "novák",
      "Novak",
      "😀",
      "",
    ]);
    expect(scan("😀 .a+b[0] Nováka")).toEqual({
      present: [".a+b[0]", "Novák", "Nováka", "😀", ""],
      missing: ["novák", "Novak"],
    });
    expect(createLiteralChecklist([])("anything")).toEqual({
      present: [],
      missing: [],
    });
    expect(createLiteralChecklist(["", "x"])("")).toEqual({
      present: [""],
      missing: ["x"],
    });
  });

  test("agrees with independent substring checks over generated vocabularies", () => {
    const fragments = [
      "",
      "a",
      "aa",
      ".a+",
      "Novák",
      "Nováka",
      "novák",
      "Novak",
      "😀",
      "e\u0301",
      "é",
    ];
    for (const left of fragments) {
      for (const right of fragments) {
        const patterns = [left, right, left + right, left];
        const scan = createLiteralChecklist(patterns);
        for (const text of [
          "",
          left + right,
          right + left,
          `(${left}) ${right}!`,
        ]) {
          expect(scan(text)).toEqual({
            present: patterns.filter((pattern) => text.includes(pattern)),
            missing: patterns.filter((pattern) => !text.includes(pattern)),
          });
        }
      }
    }
  });
});
