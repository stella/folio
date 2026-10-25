import { describe, expect, test } from "bun:test";

import {
  evaluateBudget,
  parseRaise,
  sumRaises,
  type SurfaceSize,
} from "./api-surface-budget";

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

  test("growth beyond the per-change headroom without a raise fails and identifies the required raise", () => {
    const result = evaluateBudget(
      surfaces(4_222),
      surfaces(4_273),
      surfaces(4_023),
    );

    expect(result.warnings).toEqual([]);
    expect(result.failures).toEqual([
      "docx-core.reportLines: base 4222, new 4273, budget 4023; add a file under scripts/api-surface-budget-raises/ raising docx-core.reportLines by at least 51",
    ]);
  });

  test("growth within the per-change headroom passes with a warning", () => {
    const result = evaluateBudget(
      surfaces(4_222),
      surfaces(4_272),
      surfaces(4_023),
    );

    expect(result.failures).toEqual([]);
    expect(result.warnings).toEqual([
      "docx-core.reportLines: base 4222, new 4272, budget 4023; growth 50 is within this change's allowance of 50",
    ]);
  });

  test("growth within a raise this change adds passes with a warning", () => {
    const result = evaluateBudget(
      surfaces(4_222),
      surfaces(4_322),
      surfaces(4_023),
      new Map([["docx-core", { reportLines: 100 }]]),
    );

    expect(result.failures).toEqual([]);
    expect(result.warnings).toEqual([
      "docx-core.reportLines: base 4222, new 4322, budget 4023; growth 100 is within this change's allowance of 100",
    ]);
  });

  test("growth beyond a raise this change adds fails", () => {
    const result = evaluateBudget(
      surfaces(4_222),
      surfaces(4_323),
      surfaces(4_023),
      new Map([["docx-core", { reportLines: 100 }]]),
    );

    expect(result.failures).toEqual([
      "docx-core.reportLines: base 4222, new 4323, budget 4023; add a file under scripts/api-surface-budget-raises/ raising docx-core.reportLines by at least 101",
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

describe("API surface budget raise files", () => {
  test("a raise maps published packages to positive metric increases", () => {
    const raise = parseRaise(
      { core: { declarationBytes: 2_633, declarationLines: 70 } },
      "raise.json",
    );

    expect(raise.isOk() && [...raise.value]).toEqual([
      ["core", { declarationBytes: 2_633, declarationLines: 70 }],
    ]);
  });

  test.each([
    [{ stale: { reportLines: 1 } }, "stale is not a published package"],
    [{ core: { pages: 1 } }, "core.pages is unknown"],
    [
      { core: { reportLines: 0 } },
      "core.reportLines must be a positive integer",
    ],
    [
      { core: { reportLines: 1.5 } },
      "core.reportLines must be a positive integer",
    ],
    [[], "expected an object of packages"],
  ])("a malformed raise %j is rejected", (value, detail) => {
    const raise = parseRaise(value, "raise.json");

    expect(raise.isErr() && raise.error.message).toBe(
      `Invalid raise.json: ${detail}`,
    );
  });

  test("raises sum per package and metric", () => {
    expect([
      ...sumRaises([
        new Map([["core", { declarationBytes: 10 }]]),
        new Map([
          ["core", { declarationBytes: 5, declarationLines: 2 }],
          ["cli", { reportLines: 1 }],
        ]),
      ]),
    ]).toEqual([
      ["core", { declarationBytes: 15, declarationLines: 2 }],
      ["cli", { reportLines: 1 }],
    ]);
  });
});
