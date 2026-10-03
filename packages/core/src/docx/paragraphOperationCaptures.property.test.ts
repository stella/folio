import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { applyDocumentOps, normalizeForOps } from "@stll/docx-core/ops";
import { packageDocumentArbitrary } from "../../../docx-core/src/ops/__tests__/packageOperationArbitraries";
import {
  generateOpSequence,
  serializedOpParts,
} from "../../../../scripts/lib/corpus-invariants/op-sequences";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "./rezip";
import { parseDocx } from "./parser";

test(
  "operation sequences retain parsed paragraph captures through composition",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        packageDocumentArbitrary,
        fc.integer({ min: 0, max: 0x7fffffff }),
        fc.constantFrom(" ", "\n", "\r\n", "\t"),
        async (document, seed, whitespace) => {
          const zip = await JSZip.loadAsync(await createDocx(document));
          const xml = await zip.file("word/document.xml")?.async("text");
          if (!xml) throw new TypeError("Missing document part");
          expect(xml).toContain("<w:pPr>");
          zip.file("word/document.xml", xml.replaceAll("<w:pPr>", `<w:pPr>${whitespace}`));
          const bytes = await zip.generateAsync({ type: "arraybuffer" });
          const parsed = normalizeForOps(await parseDocx(bytes, { preloadFonts: false }));
          const control = await serializedOpParts(parsed);
          // The sequence supplies the schedule; apply it to the parsed graph so
          // the byte oracle measures operation ownership, including private captures.
          const sequence = generateOpSequence(parsed, seed);
          const applied = applyDocumentOps(
            parsed,
            sequence.steps.map(({ op }) => op),
          ).unwrap();
          const restored = applyDocumentOps(applied.document, applied.inverse).unwrap();
          const saved = await serializedOpParts(restored.document);
          expect(saved.get("word/document.xml")).toEqual(control.get("word/document.xml"));
        },
      ),
      { numRuns: 12 },
    );
  },
  propertyTestTimeout(60_000),
);
