import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { schema } from "../prosemirror/schema";
import { resolveWholeStory } from "./wholeStoryRevisionResolution";

setDefaultTimeout(propertyTestTimeout(5_000));

const paragraphs = fc.array(fc.string({ minLength: 1, maxLength: 30 }), {
  minLength: 1,
  maxLength: 12,
});

for (const mode of ["accept", "reject"] as const) {
  test(`${mode} preserves unchanged stories and sibling identities`, () => {
    fc.assert(
      fc.property(paragraphs, (texts) => {
        const clean = texts.map((text) =>
          schema.node("paragraph", null, [schema.text(text, [schema.mark("bold")])]),
        );
        const doc = schema.node("doc", null, clean);
        expect(resolveWholeStory({ doc, mode, styleResolver: null }).resolved).toBe(doc);

        const changed = schema.node("paragraph", null, [
          schema.text("revision", [schema.mark("insertion")]),
        ]);
        const mixed = schema.node("doc", null, [...clean, changed, ...clean]);
        const resolved = resolveWholeStory({ doc: mixed, mode, styleResolver: null }).resolved;
        expect(resolved).not.toBe(mixed);
        expect(resolved.child(clean.length).textContent).toBe(mode === "accept" ? "revision" : "");
        for (const [index, paragraph] of clean.entries()) {
          expect(resolved.child(index)).toBe(paragraph);
          expect(resolved.child(clean.length + 1 + index)).toBe(paragraph);
        }
      }),
      propertyConfig(),
    );
  });
}
