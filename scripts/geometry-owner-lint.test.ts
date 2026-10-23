import { expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import { z } from "zod";

import { IMAGE_RESIZE_ADAPTERS } from "../.oxlint-plugins/folio-geometry-owners";
import { IMAGE_RESIZE_HANDLES } from "../packages/core/src/utils/imageResize";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const RULE = "folio-geometry-owners(image-resize-owner)";
const OWNER = "@stll/folio-core/utils/imageResize";
const OPTIONS =
  '{ handle: "se", deltaX: -100, deltaY: -10, startWidth: 200, startHeight: 100, aspect: "locked" }';
const IMPORT = `import { calculateResizedImageDimensions as resize } from "${OWNER}";\n`;
const CASES = {
  direct: { prelude: IMPORT, body: `return resize(${OPTIONS});`, count: 0 },
  namespace: {
    prelude: `import * as geometry from "${OWNER}";`,
    body: `return geometry["calculateResizedImageDimensions"](${OPTIONS});`,
    count: 0,
    other: ["import(namespace)"],
  },
  aliasChain: {
    prelude: `${IMPORT}\nconst calculate = resize; const dimensions = calculate;`,
    body: `return dimensions(${OPTIONS});`,
    count: 0,
  },
  destructuredNamespace: {
    prelude: `import * as geometry from "${OWNER}";\n\nconst { calculateResizedImageDimensions: resize } = geometry;`,
    body: `return resize(${OPTIONS});`,
    count: 0,
  },
  rounding: {
    prelude: IMPORT,
    body: `const size = resize(${OPTIONS}); return { width: Math.round(size.width), height: Math.round(size.height) };`,
    count: 0,
  },
  noOwner: {
    prelude: "",
    body: "return { width: 180, height: 90 };",
    count: 1,
  },
  unusedOwner: {
    prelude: `${IMPORT}\nvoid resize;`,
    body: "return { width: 180, height: 90 };",
    count: 1,
  },
  differentHandler: {
    prelude: `${IMPORT}\nconst elsewhere = () => resize(${OPTIONS}); void elsewhere;`,
    body: "return { width: 180, height: 90 };",
    count: 1,
  },
  typeOnly: {
    prelude: `import type { calculateResizedImageDimensions as resize } from "${OWNER}";`,
    body: `return resize(${OPTIONS});`,
    count: 1,
    other: (file: string) =>
      file.endsWith(".vue") ? [] : ["eslint(no-unused-vars)"],
  },
  wrongModule: {
    prelude:
      'import { calculateResizedImageDimensions as resize } from "./privateResizeMath";',
    body: `return resize(${OPTIONS});`,
    count: 1,
  },
  localMaximum: {
    prelude: IMPORT,
    body: `const size = resize(${OPTIONS}); return { ...size, width: Math.max(20, size.width) };`,
    count: 1,
  },
  computedMinimum: {
    prelude: IMPORT,
    body: `const size = resize(${OPTIONS}); return { ...size, width: Math["min"](2000, size.width) };`,
    count: 1,
  },
  boundsAlias: {
    prelude: `${IMPORT}\nconst math = globalThis.Math; const clamp = math["max"];`,
    body: `const size = resize(${OPTIONS}); return { ...size, width: clamp(20, size.width) };`,
    count: 1,
  },
  destructuredBounds: {
    prelude: `${IMPORT}\nconst { min: clamp } = Math;`,
    body: `const size = resize(${OPTIONS}); return { ...size, width: clamp(2000, size.width) };`,
    count: 1,
  },
  conditionalBounds: {
    prelude: IMPORT,
    body: `const size = resize(${OPTIONS}); return { ...size, width: size.width < 20 ? 20 : size.width };`,
    count: 1,
  },
  assignedBounds: {
    prelude: IMPORT,
    body: `const size = resize(${OPTIONS}); if (size.width < 20) size.width = 20; return size;`,
    count: 1,
  },
  shadowedOwner: {
    prelude: `${IMPORT}\nvoid resize;`,
    body: `const resize = () => ({ width: 180, height: 90 }); return resize(${OPTIONS});`,
    count: 1,
    other: ["eslint(no-shadow)"],
  },
  shadowedNamespace: {
    prelude: `import * as geometry from "${OWNER}";\n\nvoid geometry;`,
    body: `const geometry = { calculateResizedImageDimensions: () => ({ width: 180, height: 90 }) }; return geometry.calculateResizedImageDimensions(${OPTIONS});`,
    count: 1,
    other: ["eslint(no-shadow)"],
  },
  reassignedAlias: {
    prelude: `${IMPORT}\nlet calculate = resize; calculate = () => ({ width: 180, height: 90 });`,
    body: `return calculate(${OPTIONS});`,
    count: 1,
  },
  duplicateHandleDomain: {
    prelude: `${IMPORT}\ntype ResizeHandle = ${IMAGE_RESIZE_HANDLES.map((handle) => JSON.stringify(handle)).join(" | ")};\nconst handle: ResizeHandle = "se"; void handle;`,
    body: `return resize(${OPTIONS});`,
    count: 1,
  },
};

test("image resize ownership covers both production adapter paths and bypass forms", () => {
  const fixtureRoot = path.join(
    REPOSITORY_ROOT,
    "test",
    "__fixtures__",
    "geometry-owner",
  );
  mkdirSync(fixtureRoot, { recursive: true });
  const directory = mkdtempSync(path.join(fixtureRoot, "run-"));
  const fixtures: Array<{ filename: string; count: number; other: string[] }> =
    [];

  try {
    for (const [adapterName, adapter] of Object.entries(
      IMAGE_RESIZE_ADAPTERS,
    )) {
      for (const [caseName, fixture] of Object.entries(CASES)) {
        const { prelude, body, count } = fixture;
        const filename = path.join(
          directory,
          `${adapterName}-${caseName}`,
          adapter.file,
        );
        const script = `${prelude}\n\nconst ${adapter.handler} = () => { ${body} };\nvoid ${adapter.handler};`;
        const source = adapter.file.endsWith(".vue")
          ? `<script setup lang="ts">\n${script}\n</script>`
          : script;
        mkdirSync(path.dirname(filename), { recursive: true });
        writeFileSync(filename, source);
        const other = "other" in fixture ? fixture.other : [];
        fixtures.push({
          filename,
          count,
          other: typeof other === "function" ? other(adapter.file) : other,
        });
      }
    }

    const result = Bun.spawnSync(
      [
        "bun",
        "--bun",
        "oxlint",
        "-c",
        "oxlint.fixtures.config.ts",
        "--no-ignore",
        "--format",
        "json",
        directory,
      ],
      { cwd: REPOSITORY_ROOT },
    );
    const output = `${result.stdout.toString()}${result.stderr.toString()}`;
    expect(output).not.toContain("Failed to load");
    expect(output).not.toContain("Failed to parse");
    const { diagnostics } = z
      .object({
        diagnostics: z.array(
          z.object({ code: z.string(), filename: z.string() }),
        ),
      })
      .parse(JSON.parse(result.stdout.toString()));

    for (const fixture of fixtures) {
      const codes = diagnostics
        .filter(
          ({ filename }) =>
            path.resolve(REPOSITORY_ROOT, filename) === fixture.filename,
        )
        .map(({ code }) => code);
      expect(
        codes.filter((code) => code === RULE),
        fixture.filename,
      ).toHaveLength(fixture.count);
      expect(
        codes.filter((code) => code !== RULE).sort(),
        fixture.filename,
      ).toEqual([...fixture.other].sort());
    }
    expect(result.exitCode).toBe(1);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}, 30_000);
