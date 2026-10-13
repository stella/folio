import { expect, test } from "bun:test";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { declarationResolutionDiagnostics } from "./lib/packed-declaration-diagnostics";

// Production-shaped consumer: editor declares model, but only model's root is exported.
test("published declarations reject unresolved subpaths of declared dependencies", async () => {
  const consumer = await mkdtemp(path.join(tmpdir(), "folio-declaration-consumer-"));
  try {
    const editor = path.join(consumer, "node_modules", "@example", "editor");
    const model = path.join(consumer, "node_modules", "model");
    await mkdir(editor, { recursive: true });
    await mkdir(model, { recursive: true });
    await writeFile(
      path.join(editor, "package.json"),
      JSON.stringify({
        name: "@example/editor",
        dependencies: { model: "1.0.0" },
        exports: { ".": { types: "./index.d.ts" } },
      }),
    );
    await writeFile(
      path.join(model, "package.json"),
      JSON.stringify({
        name: "model",
        exports: { ".": { types: "./index.d.ts" } },
      }),
    );
    await writeFile(path.join(model, "index.d.ts"), "export type Model = { value: string };\n");
    await writeFile(
      path.join(consumer, "consumer.ts"),
      'import type { Model } from "@example/editor"; export type Consumer = Model;',
    );
    const projectFile = path.join(consumer, "tsconfig.json");
    await writeFile(
      projectFile,
      JSON.stringify({
        compilerOptions: {
          module: "Preserve",
          moduleResolution: "Bundler",
          strict: true,
          noEmit: true,
          skipLibCheck: true,
          types: [],
        },
        files: ["consumer.ts"],
      }),
    );
    const check = () =>
      declarationResolutionDiagnostics({ projectFile, packageNames: ["@example/editor"] });
    await writeFile(path.join(editor, "index.d.ts"), 'export type { Model } from "model/private";');
    expect(check().map(({ code }) => code)).toContain(2307);
    await writeFile(path.join(editor, "index.d.ts"), 'export type { Missing } from "model";');
    expect(check().map(({ code }) => code)).toContain(2305);
    await writeFile(path.join(editor, "index.d.ts"), 'export type { Model } from "model";');
    expect(check()).toEqual([]);
    await writeFile(
      path.join(editor, "index.d.ts"),
      'import type { Buffer } from "node:buffer"; export type Model = Buffer;',
    );
    expect(check()).toEqual([]);
    await writeFile(path.join(editor, "index.d.ts"), 'export type { Model } from "model";');
    // A third-party declaration error is not attributed to editor.
    await writeFile(
      path.join(model, "index.d.ts"),
      'export type Model = import("third-party/private").Model;',
    );
    expect(check()).toEqual([]);
  } finally {
    await rm(consumer, { recursive: true, force: true });
  }
});
