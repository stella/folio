import { describe, expect, test } from "bun:test";

import { evaluateBudget, type SurfaceSize } from "./api-surface-budget";

const size = (reportLines: number): SurfaceSize => ({
  reportBytes: 100,
  reportLines,
  declarationBytes: 100,
  declarationLines: 10,
});

const surfaces = (reportLines: number) =>
  new Map([["docx-core", size(reportLines)]]);

describe("API surface budget comparison", () => {
  test("a new package with a committed budget passes", () => {
    expect(evaluateBudget(new Map(), surfaces(4_223), surfaces(4_223))).toEqual(
      {
        failures: [],
        warnings: [],
      },
    );
  });

  test("a new package without a committed budget fails and names the package", () => {
    expect(evaluateBudget(new Map(), surfaces(4_223), new Map())).toEqual({
      failures: ["Missing docx-core API surface budget"],
      warnings: [],
    });
  });

  test("growth without a budget bump fails and identifies the required bump", () => {
    const result = evaluateBudget(
      surfaces(4_222),
      surfaces(4_223),
      surfaces(4_023),
    );

    expect(result.warnings).toEqual([]);
    expect(result.failures).toEqual([
      "docx-core.reportLines: base 4222, new 4223, budget 4023; bump scripts/api-surface-budget.json docx-core.reportLines to at least 4223",
    ]);
  });

  test("growth with an exact budget bump passes", () => {
    expect(
      evaluateBudget(surfaces(4_222), surfaces(4_223), surfaces(4_223)),
    ).toEqual({
      failures: [],
      warnings: [],
    });
  });

  test("no growth over a drifted base passes with a warning", () => {
    const result = evaluateBudget(
      surfaces(4_222),
      surfaces(4_222),
      surfaces(4_023),
    );

    expect(result.failures).toEqual([]);
    expect(result.warnings).toEqual([
      "docx-core.reportLines: base 4222, new 4222, budget 4023; existing base drift did not grow",
    ]);
  });
});
