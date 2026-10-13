import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { assertProperty } from "../../../../../test/property-testing";
import { createEmptyDocument } from "../../utils/createDocument";
import { buildSections } from "../../docx/documentParser";
import type { Paragraph } from "../../types/document";
import { toProseDoc } from "./toProseDoc";
import { fromProseDoc } from "./fromProseDoc";

// Save-based oracles previously hid stale section views by rebuilding the package.
test("captured sections partition the current body after edits", async () => {
  await assertProperty(
    fc.property(fc.array(fc.boolean(), { minLength: 1, maxLength: 4 }), (boundaries) => {
      const source = createEmptyDocument();
      const final = source.package.document.finalSectionProperties;
      if (!final) throw new Error("Expected final section properties");
      source.package.document.content = boundaries.map(
        (boundary, index) =>
          ({
            type: "paragraph",
            paraId: (0xa1000000 + index).toString(16).toUpperCase(),
            content: [{ type: "run", content: [{ type: "text", text: "Alpha" }] }],
            ...(boundary ? { sectionProperties: { ...final, pageWidth: 12000 + index } } : {}),
          }) satisfies Paragraph,
      );
      source.package.document.sections = buildSections(source.package.document.content, final);
      const state = EditorState.create({ doc: toProseDoc(source) });
      const changed = state.apply(state.tr.insertText(" updated", 3));
      const captured = fromProseDoc(changed.doc, source, { stylesheetSource: { type: "package" } });
      const body = captured.package.document;
      expect(body.sections?.flatMap(({ content }) => content)).toEqual(body.content);
      let offset = 0;
      for (const section of body.sections ?? []) {
        expect(section.content).toEqual(
          body.content.slice(offset, offset + section.content.length),
        );
        offset += section.content.length;
        const endpoint = section.content.at(-1);
        expect(section.properties).toEqual(
          endpoint?.type === "paragraph" && endpoint.sectionProperties
            ? endpoint.sectionProperties
            : final,
        );
      }
      expect(offset).toBe(body.content.length);
      expect(source.package.document.content.at(0)?.type).toBe("paragraph");
      expect(source.package.document.sections?.flatMap(({ content }) => content)).not.toEqual(
        body.content,
      );
    }),
    { numRuns: 12, seed: 1641003, id: "captured section views partition the edited body" },
  );
});
