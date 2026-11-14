import { describe, expect, test } from "bun:test";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { subprocessTestTimeout } from "../test/subprocess";
import {
  evaluateBudget,
  parseRaise,
  resolveBudgetBase,
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

describe("API surface budget base", () => {
  // A clone whose main moved on after the checked-out branch left it.
  const branchFixture = () => {
    const root = mkdtempSync(path.join(tmpdir(), "folio-api-budget-base-"));
    const git = (...args: string[]) => {
      const result = spawnSync(
        "git",
        [
          "-c",
          "user.name=Fixture",
          "-c",
          "user.email=fixture@example.test",
          "-c",
          "commit.gpgsign=false",
          ...args,
        ],
        { cwd: root, encoding: "utf8" },
      );
      if (result.status !== 0) throw new Error(result.stderr);
      return result.stdout.trim();
    };
    const commit = (content: string) => {
      writeFileSync(path.join(root, "file"), content);
      git("add", "file");
      git("commit", "--quiet", "-m", content);
      return git("rev-parse", "HEAD");
    };
    git("init", "--quiet", "--initial-branch=main");
    const forkPoint = commit("one");
    git("switch", "--quiet", "-c", "topic");
    commit("topic");
    git("switch", "--quiet", "main");
    git("update-ref", "refs/remotes/origin/main", commit("two"));
    git("switch", "--quiet", "topic");
    return { root, forkPoint };
  };

  test(
    "a hand-started run compares with where its branch left main",
    () => {
      const { root, forkPoint } = branchFixture();
      try {
        // The workflow passes an empty API_BASE when the event names none.
        for (const environment of [{ API_BASE: "" }, {}])
          expect(
            resolveBudgetBase({ argument: undefined, environment, root }),
          ).toBe(forkPoint);
      } finally {
        rmSync(root, { recursive: true, force: true });
      }
    },
    subprocessTestTimeout(16),
  );

  test(
    "a base the run names wins over the resolved one",
    () => {
      const { root } = branchFixture();
      const named = "0123456789abcdef0123456789abcdef01234567";
      const other = "89abcdef0123456789abcdef0123456789abcdef";
      try {
        expect(
          resolveBudgetBase({
            argument: undefined,
            environment: { API_BASE: named },
            root,
          }),
        ).toBe(named);
        expect(
          resolveBudgetBase({
            argument: named,
            environment: { API_BASE: other },
            root,
          }),
        ).toBe(named);
      } finally {
        rmSync(root, { recursive: true, force: true });
      }
    },
    subprocessTestTimeout(14),
  );
});
