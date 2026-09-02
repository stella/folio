/** Wiring tests for controller, engine, and projection boundary lint rules. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const CONTROLLER_RULE_MARKER = "folio-layer-boundaries(controller-and-engine-seams)";
const PROJECTION_RULE_MARKER = "folio-layer-boundaries(rust-projection-boundary)";

setDefaultTimeout(30_000);

const lintPath = (relativePath: string, marker: string): number => {
  const result = Bun.spawnSync(
    ["bun", "--bun", "oxlint", "-c", "oxlint.config.ts", "--no-ignore", relativePath],
    { cwd: REPOSITORY_ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  const violationCount = output.split(marker).length - 1;
  const hasViolations = violationCount > 0;
  const lintFailed = result.exitCode !== 0;
  if (hasViolations !== lintFailed) {
    throw new Error(
      `oxlint exited with ${result.exitCode} for ${violationCount} ${marker} violations:\n${output}`,
    );
  }
  return violationCount;
};

describe("boundary lint fixture runner", () => {
  test("surfaces lint invocation failures", () => {
    expect(() =>
      lintPath("test/__fixtures__/missing-boundary-fixture.ts", CONTROLLER_RULE_MARKER),
    ).toThrow("oxlint exited with 1 for 0");
  });
});

describe("controller and engine boundary lint", () => {
  test("rejects a runtime dependency from the pure layout engine", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/core/src/layout-engine/runtime-import.invalid.ts",
        CONTROLLER_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("rejects a package-self runtime dependency from the pure layout engine", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/core/src/layout-engine/package-runtime-import.invalid.ts",
        CONTROLLER_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("rejects a CommonJS runtime dependency from the pure layout engine", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/core/src/layout-engine/runtime-require.invalid.ts",
        CONTROLLER_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("rejects a controller implementation imported upstream", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/core/src/render-dom/controller-import.invalid.ts",
        CONTROLLER_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("rejects a CommonJS controller dependency imported upstream", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/core/src/render-dom/controller-require.invalid.ts",
        CONTROLLER_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("accepts a pure model dependency from the layout engine", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/core/src/layout-engine/model-import.valid.ts",
        CONTROLLER_RULE_MARKER,
      ),
    ).toBe(0);
  });

  test("accepts a CommonJS pure model dependency from the layout engine", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/core/src/layout-engine/model-require.valid.ts",
        CONTROLLER_RULE_MARKER,
      ),
    ).toBe(0);
  });
});

describe("Rust projection boundary lint", () => {
  test("rejects a TypeScript archive fallback in the projection boundary", () => {
    expect(
      lintPath("test/__fixtures__/packages/docx-core/src/projection.ts", PROJECTION_RULE_MARKER),
    ).toBe(1);
  });

  test("rejects a second generated-kernel entry point", () => {
    expect(
      lintPath("test/__fixtures__/packages/docx-core/src/alternate.ts", PROJECTION_RULE_MARKER),
    ).toBe(1);
  });

  test("rejects a CommonJS generated-kernel entry point", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/docx-core/src/alternate-require.invalid.ts",
        PROJECTION_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("rejects a direct generated-kernel WASM asset entry point", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/docx-core/src/alternate-wasm.invalid.ts",
        PROJECTION_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("rejects an opaque generated-kernel WASM asset entry point", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/docx-core/src/alternate-wasm-computed.invalid.ts",
        PROJECTION_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("rejects an opaque dynamic-import entry point", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/docx-core/src/alternate-dynamic-import.invalid.ts",
        PROJECTION_RULE_MARKER,
      ),
    ).toBe(1);
  });

  test("accepts the canonical projection boundary", () => {
    expect(lintPath("packages/docx-core/src/projection.ts", PROJECTION_RULE_MARKER)).toBe(0);
  });

  test("accepts absolute URL parsing that carries no module base", () => {
    expect(
      lintPath(
        "test/__fixtures__/packages/docx-core/src/url-parse.valid.ts",
        PROJECTION_RULE_MARKER,
      ),
    ).toBe(0);
  });

  test("accepts allow-listed CommonJS imports at the canonical projection boundary", () => {
    expect(
      lintPath(
        "test/__fixtures__/projection-require-valid/packages/docx-core/src/projection.ts",
        PROJECTION_RULE_MARKER,
      ),
    ).toBe(0);
  });

  test("accepts a projection test that loads the generated fixture", () => {
    expect(lintPath("packages/docx-core/src/projection.test.ts", PROJECTION_RULE_MARKER)).toBe(0);
  });
});
