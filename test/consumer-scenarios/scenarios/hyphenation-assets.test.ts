/**
 * Hyphenation dictionaries are data files of the installed package, not
 * modules: a host may hand one over instead of letting folio read it, a
 * dictionary that cannot be read settles as a failure rather than throwing,
 * and by default every dictionary loads from beside the package wherever the
 * runtime can read files (`process.getBuiltinModule`, Node 22.3 and later).
 */

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import {
  configureHyphenationAssets,
  HYPHENATION_DICTIONARY_IDS,
  preloadHyphenationDictionaries,
} from "@stll/folio-core/hyphenation";

const NOWHERE = new URL("file:///nonexistent-folio-hyphenation/");

const packagedDictionary = (id: string): unknown =>
  JSON.parse(
    readFileSync(
      fileURLToPath(
        import.meta.resolve(`@stll/folio-core/hyphenation/${id}.json`),
      ),
      "utf8",
    ),
  );

test("a dictionary handed over loads without reading anything", async () => {
  configureHyphenationAssets({
    baseUrl: NOWHERE,
    dictionaries: { "en-gb": packagedDictionary("en-gb") },
  });
  const loaded = await preloadHyphenationDictionaries(["en-GB"]);
  assert.ok(loaded.isOk(), loaded.isErr() ? loaded.error.message : "");
});

test("a dictionary that cannot be read settles as a failure", async () => {
  configureHyphenationAssets({ baseUrl: NOWHERE });
  const loaded = await preloadHyphenationDictionaries(["cs-CZ"]);
  assert.ok(loaded.isErr());
  assert.equal(loaded.error.dictionary, "cs");
});

test("every dictionary loads from the installed package", async () => {
  configureHyphenationAssets({});
  const loaded = await preloadHyphenationDictionaries(
    HYPHENATION_DICTIONARY_IDS,
  );
  if (typeof Reflect.get(process, "getBuiltinModule") === "function") {
    assert.ok(loaded.isOk(), loaded.isErr() ? loaded.error.message : "");
  } else {
    // Without a file reader the dictionaries must be handed over.
    assert.ok(loaded.isErr());
  }
});
