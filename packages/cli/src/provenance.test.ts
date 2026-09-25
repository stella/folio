import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { spawnSync } from "node:child_process";

import { makeTempDir } from "./__tests__/fixtures";
import { ISOLATED_GIT_ENV } from "./__tests__/io";
import { resolveAuthor, resolveTransactionDate } from "./provenance";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
});

afterEach(async () => {
  await cleanup();
});

describe("resolveAuthor", () => {
  test("prefers --author, then FOLIO_AUTHOR, then git user.name", () => {
    spawnSync("git", ["init", "-q", dir], { env: { ...process.env, ...ISOLATED_GIT_ENV } });
    spawnSync("git", ["-C", dir, "config", "user.name", "Git Person"]);
    const withEnv = { ...ISOLATED_GIT_ENV, FOLIO_AUTHOR: "Env Person" };

    expect(resolveAuthor({ explicit: "Flag Person", env: withEnv, cwd: dir }).unwrap()).toBe(
      "Flag Person",
    );
    expect(resolveAuthor({ explicit: undefined, env: withEnv, cwd: dir }).unwrap()).toBe(
      "Env Person",
    );
    expect(resolveAuthor({ explicit: " ", env: ISOLATED_GIT_ENV, cwd: dir }).unwrap()).toBe(
      "Git Person",
    );
  });

  test("refuses when no author is configured anywhere", () => {
    const result = resolveAuthor({ explicit: undefined, env: ISOLATED_GIT_ENV, cwd: dir });

    expect(result.isErr() && result.error.code).toBe("author_required");
  });
});

describe("resolveTransactionDate", () => {
  test("normalizes to UTC seconds and refuses unparsable dates", () => {
    const fixed = () => new Date("2026-03-04T05:06:07.890Z");

    expect(resolveTransactionDate(undefined, fixed).unwrap()).toBe("2026-03-04T05:06:07Z");
    expect(resolveTransactionDate("2026-01-02T03:04:05+02:00").unwrap()).toBe(
      "2026-01-02T01:04:05Z",
    );
    const invalid = resolveTransactionDate("next tuesday");
    expect(invalid.isErr() && invalid.error.code).toBe("usage_error");
  });
});
