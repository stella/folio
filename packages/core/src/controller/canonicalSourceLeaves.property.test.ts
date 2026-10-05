import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, ParagraphContent, RunContent } from "../types/document";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { createCanonicalSession } from "./canonicalSession";

const omittedLeaves = [
  { type: "instrText", text: "PAGE" },
  { type: "fieldChar", charType: "begin" },
  { type: "fieldChar", charType: "separate" },
  { type: "fieldChar", charType: "end" },
] as const satisfies readonly RunContent[];

const source = (content: ParagraphContent[]): Document => ({
  package: { document: { content: [{ type: "paragraph", paraId: "12345678", content }] } },
});

test(
  "generated omitted run leaves refuse activation without throwing or mutating their source",
  () => {
    assertProperty(
      fc.property(
        fc.constantFrom("", "prefix", "😀"),
        fc.constantFrom("", "suffix", "東京"),
        (before, after) => {
          for (const leaf of omittedLeaves) {
            for (const wrapped of [false, true]) {
              const content: ParagraphContent[] = [
                {
                  type: "run",
                  content: [
                    ...(before === "" ? [] : [{ type: "text", text: before } as const]),
                    leaf,
                    ...(after === "" ? [] : [{ type: "text", text: after } as const]),
                  ],
                },
              ];
              const document = source(
                wrapped
                  ? [{ type: "hyperlink", href: "https://example.com/", children: content }]
                  : content,
              );
              const unchanged = structuredClone(document);
              const result = createCanonicalSession(document);
              expect(result.isErr()).toBe(true);
              if (result.isErr())
                expect(result.error).toMatchObject({
                  name: "CanonicalSessionError",
                  gap: CANONICAL_GAP.dispatch,
                  reason: "refused",
                  message: "The paragraph cannot be projected as plain text.",
                });
              expect(document).toEqual(unchanged);
            }
          }
        },
      ),
      { numRuns: 12 },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "generated hyphen leaves retain their strict projection and address roundtrip",
  () => {
    assertProperty(
      fc.property(
        fc.array(fc.constantFrom("softHyphen", "noBreakHyphen"), { minLength: 1, maxLength: 6 }),
        (types) => {
          const document = source([{ type: "run", content: types.map((type) => ({ type })) }]);
          const unchanged = structuredClone(document);
          const session = createCanonicalSession(document).unwrap();
          expect(session.projection.doc.textContent).toBe(
            types.map((type) => (type === "softHyphen" ? "\u00ad" : "\u2011")).join(""),
          );
          for (let offset = 0; offset <= types.length; offset++) {
            const address = { story: "main", blockId: "12345678", offset } as const;
            const position = session.projection.positionAt(address).unwrap();
            expect(session.projection.addressAt(position).unwrap()).toMatchObject(address);
          }
          expect(document).toEqual(unchanged);
        },
      ),
      { numRuns: 12 },
    );
  },
  propertyTestTimeout(30_000),
);
