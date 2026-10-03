import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { applyDocumentOps, normalizeForOps } from "@stll/docx-core/ops";
import { captureDocumentArbitrary } from "../../../../test/generators/packageOperationArbitraries";
import {
  generateOpSequence,
  serializedOpParts,
} from "../../../../scripts/lib/corpus-invariants/op-sequences";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "./rezip";
import {
  cloneDocumentWithParagraphPropertySources,
  copyParagraphPropertyCapture,
} from "./paragraphPropertySource";
import { visitDocxParagraphs } from "./paragraphTraversal";
import { parseDocx } from "./parser";
import { createEmptyDocument } from "../utils/createDocument";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";

test(
  "operation sequences retain parsed paragraph captures through composition",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        captureDocumentArbitrary,
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
          // Mutation control: emulate the old clone losing source handles while
          // preserving every public field. This oracle must fail without the fix.
          const withoutCaptures = cloneDocumentWithParagraphPropertySources(restored.document);
          let removed = 0;
          visitDocxParagraphs({ documentBody: withoutCaptures.package.document }, (paragraph) => {
            for (const key of Object.getOwnPropertySymbols(paragraph)) {
              if (key.description !== "paragraphPropertyCapture") continue;
              expect(Reflect.deleteProperty(paragraph, key)).toBe(true);
              removed++;
            }
          });
          expect(removed).toBeGreaterThan(0);
          expect(structuredClone(withoutCaptures.package.document.content)).toStrictEqual(
            structuredClone(restored.document.package.document.content),
          );
          const mutated = await serializedOpParts(withoutCaptures);
          expect(mutated.get("word/document.xml")).not.toEqual(control.get("word/document.xml"));
        },
      ),
      { numRuns: 12 },
    );
  },
  propertyTestTimeout(60_000),
);

test(
  "mixed durable and spread-derived paragraphs retain exact public content through editor save",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.boolean(), { minLength: 1, maxLength: 6 }),
        fc.oneof(
          fc.constant(null),
          fc.constant(undefined),
          fc.string({ maxLength: 5 }),
          fc.integer(),
          fc.constant({}),
        ),
        async (derive, invalidHandle) => {
          const initial = createEmptyDocument();
          initial.package.document.content = derive.map((_, index) => ({
            type: "paragraph",
            paraId: (index + 1).toString(16).toUpperCase().padStart(8, "0"),
            formatting: { alignment: "center" },
            content: [{ type: "run", content: [{ type: "text", text: `é😀 ${index}` }] }],
          }));
          const source = await parseDocx(await createDocx(initial), {
            preloadFonts: false,
            detectVariables: false,
          });
          const content = source.package.document.content.map((paragraph, index) => {
            if (paragraph.type !== "paragraph") throw new TypeError("Expected a source paragraph");
            return derive.at(index) ? { ...paragraph, content: [...paragraph.content] } : paragraph;
          });
          const document = {
            ...source,
            package: { ...source.package, document: { ...source.package.document, content } },
          };
          const rebuilt = updateDocumentContent(document, toProseDoc(document));
          const repeated = updateDocumentContent(rebuilt, toProseDoc(rebuilt));
          expect(structuredClone(repeated.package.document.content)).toStrictEqual(
            structuredClone(rebuilt.package.document.content),
          );
          for (const paragraph of source.package.document.content) {
            if (paragraph.type !== "paragraph") throw new TypeError("Expected a source paragraph");
            const detached = { ...paragraph };
            copyParagraphPropertyCapture(detached, paragraph);
            const invalid = {
              ...source,
              package: {
                ...source.package,
                document: { ...source.package.document, content: [detached] },
              },
            };
            expect(() => updateDocumentContent(invalid, toProseDoc(source))).toThrow(
              "A source-bound paragraph is missing its paragraph-property token.",
            );
          }
          const reopened = await parseDocx(await createDocx(rebuilt), {
            preloadFonts: false,
            detectVariables: false,
          });
          expect(structuredClone(reopened.package.document.content)).toStrictEqual(
            structuredClone(document.package.document.content),
          );
          for (const paragraph of content) {
            const forged = { ...paragraph };
            const key = Object.getOwnPropertySymbols(forged).find(
              (symbol) => symbol.description === "paragraphPropertyCapture",
            );
            if (!key) throw new TypeError("Expected a paragraph capture handle");
            Object.defineProperty(forged, key, { value: invalidHandle });
            const invalid = {
              ...document,
              package: {
                ...document.package,
                document: { ...document.package.document, content: [forged] },
              },
            };
            expect(() => updateDocumentContent(invalid, toProseDoc(invalid))).toThrow(
              "The source paragraph identity is invalid.",
            );
          }
        },
      ),
      { numRuns: 20 },
    );
  },
  propertyTestTimeout(60_000),
);
