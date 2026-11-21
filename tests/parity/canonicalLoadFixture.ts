import { normalizeCanonicalSource } from "../../packages/core/src/canonical/canonicalSourceNormalization";
import { prepareCanonicalDocxInput } from "../../packages/core/src/docx/canonicalSessionInput";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { normalizeForOps } from "../../packages/docx-core/src/ops/contract";
import { validateDocxFixture } from "../support/validatedDocxFixture";

/**
 * Load and compare the same canonical input, including its allocated package
 * IDs. The content is the body as the session reads it: a body that ends in a
 * table holds the closing paragraph the session gives it.
 */
export const canonicalLoadFixture = async (buffer: ArrayBuffer) => {
  await validateDocxFixture(new Uint8Array(buffer), "canonical-load");
  const prepared = (await prepareCanonicalDocxInput(buffer)).unwrap();
  const bytes = new Uint8Array(prepared);
  const parsed = await parseDocx(bytes, {
    preloadFonts: false,
    detectVariables: false,
  });
  const read = normalizeCanonicalSource(normalizeForOps(parsed)).document;
  return {
    bytes: [...bytes],
    content: JSON.stringify(read.package.document.content),
  };
};
