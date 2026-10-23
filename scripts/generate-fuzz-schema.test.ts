import { expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { generateFuzzSchema } from "./generate-fuzz-schema";
import type { OoxmlSchemaGraph } from "./generate-ooxml-schema-graph";

const ROOT = path.resolve(import.meta.dir, "..");

test("committed fuzz attribute facts equal their schema derivation", async () => {
  const graph: OoxmlSchemaGraph = JSON.parse(
    await readFile(
      path.join(
        ROOT,
        "specifications/generated/docx-transitional-schema.gen.json",
      ),
      "utf8",
    ),
  );
  expect(generateFuzzSchema(graph)).toBe(
    await readFile(
      path.join(
        ROOT,
        "packages/docx-core/src/validate/schemaAttributes.gen.ts",
      ),
      "utf8",
    ),
  );
});
