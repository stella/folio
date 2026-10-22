import { expect, test } from "bun:test";

import { shapeArrayBuffer } from "../packages/core/src/__tests__/documentShapes";
import { createCanonicalSession } from "../packages/core/src/controller/canonicalSession";
import { cloneDocumentWithParagraphPropertySources } from "../packages/core/src/docx/paragraphPropertySource";
import { parseDocx } from "../packages/core/src/docx/parser";
import { toProseDoc } from "../packages/core/src/prosemirror/conversion/toProseDoc";
import {
  canonicalBrowserAcceptances,
  canonicalBrowserSourceRefusals,
} from "../tests/visual/canonical-browser-acceptance-traces";

for (const { seed, trace } of canonicalBrowserAcceptances) {
  test(`acceptance seed ${seed} classifies its canonical source fixture`, async () => {
    const document = await parseDocx(await shapeArrayBuffer(trace.shape), {
      preloadFonts: false,
      detectVariables: false,
    });
    const original = cloneDocumentWithParagraphPropertySources(document);
    const result = createCanonicalSession(document);
    // Temporary table (#1474) and image source limits follow canonical support;
    // the same traces become accepted without a permanent refusal guard.
    if (result.isErr()) {
      expect(result.error).toMatchObject(
        canonicalBrowserSourceRefusals[trace.shape],
      );
      expect(document).toEqual(original);
    } else
      expect(
        result.value.projection.doc.eq(toProseDoc(result.value.document)),
      ).toBe(true);
    expect(trace.actions.length).toBeGreaterThan(0);
  });
}
