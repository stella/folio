/** Wiring tests for the derived identity-attribute set, its lint rule, and its baseline. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import { IDENTITY_ATTRIBUTES } from "../specifications/generated/identity-attributes.gen";
import baseline from "./identity-attribute-baseline.json";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-identity-attributes(no-prefix-resolved-identity-read)";

setDefaultTimeout(30_000);

const lintFixture = (fixture: string): number => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.config.ts",
      "--no-ignore",
      path.join("test", "__fixtures__", fixture),
    ],
    { cwd: REPO_ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  return output.split(RULE_MARKER).length - 1;
};

describe("no-prefix-resolved-identity-read", () => {
  test("flags every prefix-resolved read of a derived identity attribute", () => {
    expect(lintFixture("identity-attributes.invalid.ts")).toBe(8);
  });

  test("accepts a namespace-resolved read, an unprefixed read, and a same-named neighbour", () => {
    expect(lintFixture("identity-attributes.valid.ts")).toBe(0);
  });
});

describe("identity-attribute set", () => {
  test("covers the identities the parsers join on", () => {
    const missing = [
      "durableId",
      "embed",
      "id",
      "Ignorable",
      "link",
      "name",
      "paraId",
      "paraIdParent",
      "space",
      "textId",
    ].filter((attribute) => !IDENTITY_ATTRIBUTES.has(attribute));
    expect(missing).toEqual([]);
  });

  test("keeps a relationship reference scoped to the relationship prefix", () => {
    expect(IDENTITY_ATTRIBUTES.get("embed")?.prefixes).toEqual(["r"]);
    expect(IDENTITY_ATTRIBUTES.get("id")?.prefixes).toEqual(["r", "w"]);
  });

  test("claims every prefix only for a local name no vocabulary shares", () => {
    // `paraId` is a Word extension; `space` is also `w:cols/@w:space`, so it
    // names the one prefix it means.
    expect(IDENTITY_ATTRIBUTES.get("paraId")?.prefixes).toBeNull();
    expect(IDENTITY_ATTRIBUTES.get("space")?.prefixes).toEqual(["xml"]);
  });

  test("every attribute records why it carries an identity", () => {
    const empty = [...IDENTITY_ATTRIBUTES.entries()]
      .filter(([, { reason }]) => reason.trim().length < 20)
      .map(([attribute]) => attribute);
    expect(empty).toEqual([]);
  });

  test("the generated set is what the schema graph and the extension table say", () => {
    const result = Bun.spawnSync(["bun", "scripts/generate-identity-attributes.ts", "check"], {
      cwd: REPO_ROOT,
    });
    expect(`${result.stdout.toString()}${result.stderr.toString()}`).toContain("up to date");
    expect(result.exitCode).toBe(0);
  });
});

describe("identity-attribute baseline", () => {
  test("only shrinks: every baseline file still exists", async () => {
    const missing: string[] = [];
    for (const file of Object.keys(baseline)) {
      // oxlint-disable-next-line no-await-in-loop -- one stat per baseline entry
      if (!(await Bun.file(path.join(REPO_ROOT, file)).exists())) {
        missing.push(file);
      }
    }
    expect(missing).toEqual([]);
  });

  test("records a positive count per file", () => {
    const invalid = Object.entries(baseline)
      .filter(([, count]) => !Number.isInteger(count) || count < 1)
      .map(([file]) => file);
    expect(invalid).toEqual([]);
  });
});
