import { expect, test } from "bun:test";
import fc from "fast-check";

import { prepareCanonicalDocxInput } from "../../packages/core/src/docx/canonicalSessionInput";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { normalizeForOps } from "../../packages/docx-core/src/ops/contract";
import { assertProperty } from "../../test/property-testing";
import { canonicalLoadFixture } from "./canonicalLoadFixture";

const AUTHORED_ID = "12345678";
const HEX_ID = /^[0-9A-F]{8}$/u;

const identities = [
  "missing",
  "paragraph id only",
  "authored",
  "duplicate",
] as const;
type IdentityClass = (typeof identities)[number];
type WrittenIds = { readonly paraId?: string; readonly textId?: string };

/**
 * The paragraph ids each input class writes. A blank document starts with a
 * paragraph id of its own, so every class states both attributes.
 */
const IDENTITY_CLASSES: Record<
  IdentityClass,
  readonly [WrittenIds, ...WrittenIds[]]
> = {
  missing: [{}],
  "paragraph id only": [{ paraId: AUTHORED_ID }],
  authored: [{ paraId: AUTHORED_ID, textId: AUTHORED_ID }],
  duplicate: [
    { paraId: AUTHORED_ID, textId: AUTHORED_ID },
    { paraId: AUTHORED_ID, textId: AUTHORED_ID },
  ],
};

test("canonical load fixtures match the loader across paragraph identity input classes", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.string({
        unit: fc.constantFrom("a", "b", " ", "é", "😀"),
        maxLength: 20,
      }),
      async (text) => {
        for (const identity of identities) {
          const written = IDENTITY_CLASSES[identity];
          const document = createEmptyDocument({ initialText: text });
          const first = document.package.document.content.at(0);
          if (first?.type !== "paragraph")
            throw new TypeError("Expected fixture paragraph.");
          const [own, ...repeated] = written;
          delete first.paraId;
          delete first.textId;
          Object.assign(first, own);
          for (const ids of repeated)
            document.package.document.content.push({
              type: "paragraph",
              ...ids,
              content: [{ type: "run", content: [{ type: "text", text }] }],
            });
          const fixture = await canonicalLoadFixture(
            await createDocx(document),
          );
          // The real loader prepares the already-prepared browser fixture again.
          const input = (
            await prepareCanonicalDocxInput(new Uint8Array(fixture.bytes))
          ).unwrap();
          const loaded = await parseDocx(input, {
            preloadFonts: false,
            detectVariables: false,
          });
          const content = normalizeForOps(loaded).package.document.content;
          expect(JSON.stringify(content)).toBe(fixture.content);
          const paragraphs = content.flatMap((block) =>
            block.type === "paragraph" ? [block] : [],
          );
          expect(paragraphs).toHaveLength(written.length);
          const seen = new Set<string>();
          for (const [index, paragraph] of paragraphs.entries()) {
            const source: WrittenIds = written.at(index) ?? {};
            expect(paragraph.paraId).toMatch(HEX_ID);
            const paraId = paragraph.paraId ?? "";
            if (source.paraId !== undefined && !seen.has(source.paraId)) {
              // The package's own first id is kept, its text id as written.
              expect(paraId).toBe(source.paraId);
              expect(paragraph.textId).toBe(source.textId);
            } else {
              // A missing or repeated id is minted with a matching text id.
              expect(paraId).not.toBe(source.paraId);
              expect(paragraph.textId).toBe(paraId);
            }
            expect(seen.has(paraId)).toBe(false);
            seen.add(paraId);
          }
        }
      },
    ),
    {
      numRuns: 20,
      id: "canonical load fixtures match the loader across paragraph identity input classes",
    },
  );
});
