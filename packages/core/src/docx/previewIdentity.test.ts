/**
 * A preview is charged because the ledger built it, so a producer that builds
 * its own image is a preview the budget cannot charge. That is how a
 * WordprocessingGroup preview once came to have no bound at all: it wrote its
 * own data-URL prefix, mime type and filename, the budget never recognized
 * them, and no package was charged for the characters it retained.
 *
 * Catching that by review is the discipline that already failed once, so it is
 * a scan instead: the media type in a data URL is a preview's identity, and
 * `previewBudget.ts` is the only module in the package that may write one.
 * Everywhere else either builds the URL from a mime type it was given
 * (`data:${…}`, which this does not match) or asks the ledger for the image.
 */

import { expect, test } from "bun:test";
import { readdir } from "node:fs/promises";
import path from "node:path";

const PACKAGE_SOURCE_ROOT = path.resolve(import.meta.dir, "..");

/** The table's own module, which is where these strings are supposed to live. */
const TABLE_MODULE = path.join(PACKAGE_SOURCE_ROOT, "docx", "previewBudget.ts");

/** `data:` followed by a media type spelled out, rather than interpolated. */
const STATIC_MEDIA_TYPE_URL = /data:[a-z]+\/[a-z0-9.+-]+/u;

const isProductModule = (file: string): boolean =>
  (file.endsWith(".ts") || file.endsWith(".tsx")) &&
  !file.endsWith(".test.ts") &&
  !file.endsWith(".test.tsx") &&
  !file.endsWith(".d.ts") &&
  file !== TABLE_MODULE;

const productModules = async (): Promise<string[]> => {
  const entries = await readdir(PACKAGE_SOURCE_ROOT, { recursive: true });
  return entries.map((entry) => path.join(PACKAGE_SOURCE_ROOT, entry)).filter(isProductModule);
};

/** The file's code without its comments, so prose may name a media type. */
const codeOf = async (file: string): Promise<string> =>
  new Bun.Transpiler({ loader: file.endsWith(".tsx") ? "tsx" : "ts" }).transformSync(
    await Bun.file(file).text(),
  );

test("no module spells a data-URL media type the preview table should own", async () => {
  const pattern = new RegExp(STATIC_MEDIA_TYPE_URL, "gu");
  const offenders: string[] = [];
  for (const file of await productModules()) {
    for (const match of (await codeOf(file)).matchAll(pattern)) {
      offenders.push(`${path.relative(PACKAGE_SOURCE_ROOT, file)}: ${match[0]}`);
    }
  }
  expect(offenders).toEqual([]);
});
