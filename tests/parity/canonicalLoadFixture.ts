import { prepareCanonicalDocxInput } from "../../packages/core/src/docx/canonicalSessionInput";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { normalizeForOps } from "../../packages/docx-core/src/ops/contract";

/** Load and compare the same canonical input, including its allocated package IDs. */
export const canonicalLoadFixture = async (buffer: ArrayBuffer) => {
  const prepared = (await prepareCanonicalDocxInput(buffer)).unwrap();
  const bytes = new Uint8Array(prepared);
  const parsed = await parseDocx(bytes, { preloadFonts: false, detectVariables: false });
  return {
    bytes: [...bytes],
    content: JSON.stringify(normalizeForOps(parsed).package.document.content),
  };
};
