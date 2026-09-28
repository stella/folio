#!/usr/bin/env bun
/** Save one synthetic edit and its structural view for inspection. */

import { createHash } from "node:crypto";
import { mkdir } from "node:fs/promises";
import path from "node:path";

import { EDIT_OPERATION_SCRIPTS } from "./edit-operation-scripts";
import { readEditStructure } from "./edit-structure";
import { runEditOperationScript } from "./run-edit-operation-script";

const run = async (): Promise<void> => {
  const script = EDIT_OPERATION_SCRIPTS.find(({ id }) => id === process.argv[2]);
  if (!script) {
    throw new Error(`Choose one script: ${EDIT_OPERATION_SCRIPTS.map(({ id }) => id).join(", ")}`);
  }
  const outputDir = path.resolve(process.argv[3] ?? ".cache/edit-cases");
  await mkdir(outputDir, { recursive: true });
  const source = await Bun.file(path.join(import.meta.dir, script.seed)).arrayBuffer();
  const saved = await runEditOperationScript(source, script);
  const result = {
    script: script.id,
    seedSha256: createHash("sha256").update(new Uint8Array(source)).digest("hex"),
    structure: await readEditStructure(saved),
  };
  await Bun.write(path.join(outputDir, `${script.id}.docx`), saved);
  await Bun.write(
    path.join(outputDir, `${script.id}.json`),
    `${JSON.stringify(result, null, 2)}\n`,
  );
  console.log(path.join(outputDir, `${script.id}.json`));
};

if (import.meta.main) await run();
