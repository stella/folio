/** Wiring tests for the reserved-value registry, its lint rule, and its baseline. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  RESERVED_VALUE_NAMESPACE_URIS,
  type ReservedValueDisposition,
} from "../packages/docx-core/src/model/reserved/disposition";
import {
  RESERVED_VALUE_REGISTRY,
  reservedValueEntries,
} from "../packages/docx-core/src/model/reserved/registry";
import { RESERVED_VALUE_READERS } from "../packages/docx-core/src/model/reserved/readers";
import baseline from "./reserved-value-baseline.json";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-reserved-values(no-bare-reserved-compare)";

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

describe("no-bare-reserved-compare", () => {
  test("flags every bare comparison against a recorded sentinel", () => {
    expect(lintFixture("reserved-values.invalid.ts")).toBe(5);
  });

  test("accepts the owning reader, a named constant, and a foreign vocabulary", () => {
    expect(lintFixture("reserved-values.valid.ts")).toBe(0);
  });
});

describe("reserved-value registry", () => {
  const entries = reservedValueEntries();

  test("covers the model types it claims to", () => {
    expect(Object.keys(RESERVED_VALUE_REGISTRY).length).toBeGreaterThan(30);
    expect(entries.length).toBeGreaterThan(250);
  });

  test("every slot resolves to a declared namespace prefix", () => {
    const prefixes = new Set(Object.keys(RESERVED_VALUE_NAMESPACE_URIS));
    const unknown: string[] = [];
    for (const { modelType, field, disposition } of entries) {
      if (disposition === "no-reserved-value") {
        continue;
      }
      for (const slot of disposition.slot.split("|")) {
        const prefix = slot.slice(0, slot.indexOf(":"));
        if (!prefixes.has(prefix)) {
          unknown.push(`${modelType}.${field}: ${slot}`);
        }
      }
    }
    expect(unknown).toEqual([]);
  });

  test("every named reader is a real function in a real module", async () => {
    const missing: string[] = [];
    for (const reader of Object.values(RESERVED_VALUE_READERS)) {
      const [modulePath, name] = reader.split("#");
      if (modulePath === undefined || name === undefined) {
        missing.push(`${reader}: not "<module>#<name>"`);
        continue;
      }
      const file = Bun.file(path.join(REPO_ROOT, modulePath));
      if (!(await file.exists())) {
        missing.push(`${reader}: no such module`);
        continue;
      }
      const source = await file.text();
      const declares = new RegExp(
        String.raw`(?:function|const|let|class)\s+${name}\b|\b${name}\s*[:=]\s*(?:\(|function|async)`,
        "u",
      );
      if (!declares.test(source)) {
        missing.push(`${reader}: module does not declare it`);
      }
    }
    expect(missing).toEqual([]);
  });

  test("every reader-owned entry names a reader from the shared table", () => {
    const known = new Set<string>(Object.values(RESERVED_VALUE_READERS));
    const strays = entries
      .filter(
        (
          entry,
        ): entry is typeof entry & { disposition: Extract<ReservedValueDisposition, object> } =>
          entry.disposition !== "no-reserved-value",
      )
      .filter(
        ({ disposition }) =>
          disposition.disposition === "reader-owned" && !known.has(disposition.reader),
      )
      .map(({ modelType, field }) => `${modelType}.${field}`);
    expect(strays).toEqual([]);
  });

  test("every evidence id it cites has a record", async () => {
    const cited = new Set<string>();
    for (const { disposition } of entries) {
      if (disposition === "no-reserved-value") {
        continue;
      }
      if (disposition.evidence !== undefined) {
        cited.add(disposition.evidence);
      }
    }
    const missing: string[] = [];
    for (const id of [...cited].toSorted()) {
      const record = Bun.file(
        path.join(REPO_ROOT, "specifications/evidence/records", `${id}.json`),
      );
      // oxlint-disable-next-line no-await-in-loop -- one small file per cited id
      if (!(await record.exists())) {
        missing.push(id);
      }
    }
    expect(missing).toEqual([]);
  });

  test("a not-modelled entry carries a reason", () => {
    const empty = entries
      .filter(
        ({ disposition }) =>
          disposition !== "no-reserved-value" &&
          disposition.disposition === "not-modelled" &&
          disposition.reason.trim().length < 40,
      )
      .map(({ modelType, field }) => `${modelType}.${field}`);
    expect(empty).toEqual([]);
  });
});

describe("reserved-value baseline", () => {
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
