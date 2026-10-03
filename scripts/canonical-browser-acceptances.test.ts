import { cloneDocumentWithParagraphPropertySources } from "../packages/core/src/docx/paragraphPropertySource";
import { expect, test } from "bun:test";
import { canonicalBrowserAcceptances } from "../tests/visual/canonical-browser-acceptance-traces";
import { shapeArrayBuffer } from "../packages/core/src/__tests__/documentShapes";
import { parseDocx } from "../packages/core/src/docx/parser";
import { createCanonicalSession } from "../packages/core/src/controller/canonicalSession";
import { toProseDoc } from "../packages/core/src/prosemirror/conversion/toProseDoc";

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
      expect(result.error.name).toBe("CanonicalSessionError");
      expect(result.error.message).toBe(
        "Canonical sessions currently require plain paragraphs and note references without revisions.",
      );
      expect(document).toEqual(original);
    } else expect(result.value.projection.doc.eq(toProseDoc(result.value.document))).toBe(true);
    expect(trace.actions.length).toBeGreaterThan(0);
  });
}
