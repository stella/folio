/**
 * Only the capture owner serializes markup a rebuilt part replays.
 *
 * The lint rule in `.oxlint-plugins/folio-verbatim-capture.ts` enforces this on
 * every run; this walks the same import edges from the source tree, so a
 * loosened lint config cannot silently let a parser reach `elementToXml` again
 * and replay a Strict fragment under a Transitional root.
 */

import { describe, expect, test } from "bun:test";
import path from "node:path";

const DOCX_DIR = path.resolve(import.meta.dir, "..");

/** `metadataPrivacy` rewrites `docProps/core.xml` under the root that part keeps. */
const ALLOWED = new Set(["verbatimCapture.ts", "metadataPrivacy.ts"]);

const IMPORTS_ELEMENT_TO_XML = /import\s*\{[^}]*\belementToXml\b[^}]*\}\s*from\s*"[^"]*xmlParser"/u;

describe("verbatim capture ownership", () => {
  test("no parser imports the raw serializer", async () => {
    const offenders: string[] = [];
    for await (const file of new Bun.Glob("**/*.ts").scan({ absolute: true, cwd: DOCX_DIR })) {
      const name = path.basename(file);
      if (ALLOWED.has(name) || name.endsWith(".test.ts")) {
        continue;
      }
      if (IMPORTS_ELEMENT_TO_XML.test(await Bun.file(file).text())) {
        offenders.push(path.relative(DOCX_DIR, file));
      }
    }
    expect(offenders).toEqual([]);
  });
});
