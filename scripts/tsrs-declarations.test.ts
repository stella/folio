import { expect, test } from "bun:test";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { build } from "tsdown";

import { scanTestTimeout } from "../test/scan-timeout";
import { tsrsDeclarations } from "./lib/tsrs-declarations";

test(
  "native declaration generation retains published union documentation",
  async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "folio-tsrs-declarations-"),
    );
    try {
      await writeFile(
        path.join(directory, "package.json"),
        JSON.stringify({ type: "module" }),
      );
      await writeFile(
        path.join(directory, "entry.ts"),
        `
/** Published operation outcome. */
export type Outcome =
  /** The operation completed. */
  | { kind: "completed" }
  /** The operation was refused. */
  | { kind: "refused"; reason: string };
`,
      );
      await writeFile(
        path.join(directory, "tsconfig.json"),
        JSON.stringify({
          compilerOptions: {
            module: "Preserve",
            moduleResolution: "Bundler",
            declaration: true,
            emitDeclarationOnly: true,
            types: [],
          },
          files: ["entry.ts"],
        }),
      );
      await build({
        cwd: directory,
        config: false,
        entry: ["entry.ts"],
        outDir: "dist",
        platform: "neutral",
        unbundle: true,
        dts: { ...tsrsDeclarations("tsconfig.json"), dtsOnly: true },
      });
      const declaration = await readFile(
        path.join(directory, "dist/entry.d.ts"),
        "utf8",
      );
      expect(declaration).toContain("Published operation outcome.");
      expect(declaration).toContain("The operation completed.");
      expect(declaration).toContain("The operation was refused.");
      expect(declaration).toContain('kind: "refused"');
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  },
  scanTestTimeout(30_000),
);
