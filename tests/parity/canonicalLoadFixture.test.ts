import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty } from "../../test/property-testing";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { prepareCanonicalDocxInput } from "../../packages/core/src/docx/canonicalSessionInput";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { normalizeForOps } from "../../packages/docx-core/src/ops/contract";
import { canonicalLoadFixture } from "./canonicalLoadFixture";

test("canonical load fixtures match the loader across paragraph identity input classes", async () => {
  let cases = 0;
  const identities = ["missing", "authored", "duplicate"] as const;
  const exercised = new Set<(typeof identities)[number]>();
  await assertProperty(
    fc.asyncProperty(
      fc.string({ unit: fc.constantFrom("a", "b", " ", "é", "😀"), maxLength: 20 }),
      async (text) => {
        for (const identity of identities) {
          cases += 1;
          exercised.add(identity);
          const document = createEmptyDocument({ initialText: text });
          const first = document.package.document.content.at(0);
          if (first?.type !== "paragraph") throw new TypeError("Expected fixture paragraph.");
          if (identity !== "missing") {
            first.paraId = "12345678";
            first.textId = "12345678";
          }
          if (identity === "duplicate")
            document.package.document.content.push({
              type: "paragraph",
              paraId: "12345678",
              textId: "12345678",
              content: [{ type: "run", content: [{ type: "text", text }] }],
            });
          const fixture = await canonicalLoadFixture(await createDocx(document));
          // The real loader prepares the already-prepared browser fixture again.
          const input = (await prepareCanonicalDocxInput(new Uint8Array(fixture.bytes))).unwrap();
          const loaded = await parseDocx(input, { preloadFonts: false, detectVariables: false });
          const content = normalizeForOps(loaded).package.document.content;
          expect(JSON.stringify(content)).toBe(fixture.content);
          const ids = content.flatMap((paragraph) => {
            if (paragraph.type !== "paragraph") return [];
            expect(paragraph.paraId).toMatch(/^[0-9A-F]{8}$/u);
            expect(paragraph.textId).toMatch(/^[0-9A-F]{8}$/u);
            return [paragraph.paraId];
          });
          expect(new Set(ids).size).toBe(ids.length);
        }
      },
    ),
    {
      numRuns: 20,
      id: "canonical load fixtures match the loader across paragraph identity input classes",
    },
  );
  expect(cases).toBeGreaterThan(0);
  expect([...exercised].sort()).toEqual([...identities].sort());
});
