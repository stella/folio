import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { scanDistUrlTargets } from "./dist-url-targets";

// A miniature published dist tree: the worker target exists as emitted `.js`;
// no `.ts` ships (the build renames source extensions).
let distDir: string;

beforeAll(async () => {
  distDir = await mkdtemp(path.join(tmpdir(), "folio-url-scan-"));
  await mkdir(path.join(distDir, "layout-engine", "measure"), { recursive: true });
  await writeFile(
    path.join(distDir, "layout-engine", "measure", "font-metrics.worker.js"),
    "export {};\n",
  );
});

afterAll(async () => {
  await rm(distDir, { recursive: true, force: true });
});

const measureModule = (code: string) => [{ file: "layout-engine/measure/measureWorker.js", code }];

describe("scanDistUrlTargets", () => {
  test("an emitted .js target resolves", () => {
    const scan = scanDistUrlTargets(
      measureModule('new Worker(new URL("font-metrics.worker.js", import.meta.url));'),
      distDir,
    );
    expect(scan.total).toBe(1);
    expect(scan.dangling).toEqual([]);
  });

  test("a ?query-suffixed target passes when the file exists", () => {
    const scan = scanDistUrlTargets(
      measureModule('new Worker(new URL("./font-metrics.worker.js?worker", import.meta.url));'),
      distDir,
    );
    expect(scan.total).toBe(1);
    expect(scan.dangling).toEqual([]);
  });

  test("a #hash-suffixed target passes when the file exists", () => {
    const scan = scanDistUrlTargets(
      measureModule('new URL("./font-metrics.worker.js#frag", import.meta.url);'),
      distDir,
    );
    expect(scan.dangling).toEqual([]);
  });

  test("a dangling .ts target is reported (the 0.1.1 regression)", () => {
    const scan = scanDistUrlTargets(
      measureModule('new Worker(new URL("font-metrics.worker.ts", import.meta.url));'),
      distDir,
    );
    expect(scan.total).toBe(1);
    expect(scan.dangling).toEqual([
      "layout-engine/measure/measureWorker.js -> font-metrics.worker.ts",
    ]);
  });

  test("a query-suffixed target whose file is missing is still reported", () => {
    const scan = scanDistUrlTargets(
      measureModule('new URL("./missing.worker.js?worker", import.meta.url);'),
      distDir,
    );
    expect(scan.dangling).toEqual([
      "layout-engine/measure/measureWorker.js -> ./missing.worker.js?worker",
    ]);
  });

  test("absolute and protocol-relative URLs are out of scope", () => {
    const scan = scanDistUrlTargets(
      measureModule(
        'new URL("https://example.com/x.ts", import.meta.url);\n' +
          'new URL("data:text/javascript,export{}", import.meta.url);\n' +
          'new URL("//cdn.example.com/y.js", import.meta.url);',
      ),
      distDir,
    );
    expect(scan.total).toBe(0);
    expect(scan.dangling).toEqual([]);
  });

  test("a parent-relative target resolves against the referencing module", () => {
    const scan = scanDistUrlTargets(
      [
        {
          file: "layout-engine/measure/deep/nested.js",
          code: 'new URL("../font-metrics.worker.js", import.meta.url);',
        },
      ],
      distDir,
    );
    expect(scan.dangling).toEqual([]);
  });
});
