import { expect, setDefaultTimeout, test } from "bun:test";

setDefaultTimeout(propertyTestTimeout(30_000));
import fc from "fast-check";
import { paragraphNumberingReference } from "@stll/docx-core/model";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { listRenderingAttrPatch } from "./listRenderingAttrs";
import { CLEARED_LIST_RENDERING_ATTRS } from "./listMarker";
import { schema } from "./schema";

test("folded LISTNUM display text stays literal while numbered and symbol-font bullet markers resolve", () => {
  assertProperty(
    fc.property(
      fc.record({
        suffix: fc.string({
          unit: fc.constantFrom("a", "ž", "😀", "%", "1", "\t"),
          minLength: 1,
          maxLength: 12,
        }),
        start: fc.integer({ min: 1, max: 9 }),
        kind: fc.constantFrom("decimal", "bullet"),
        templateKind: fc.constantFrom("dynamic", "literal", "cleared"),
      }),
      ({ suffix, start, kind, templateKind }) => {
        const bullet = kind === "bullet";
        const decimalTemplate = templateKind === "literal" ? "Article" : "%1.";
        const decimalMarker = templateKind === "literal" ? "Article" : `${start}.`;
        const nextDecimalMarker = templateKind === "literal" ? "Article" : `${start + 1}.`;
        const markerTemplate = bullet ? "\uF0B7" : decimalTemplate;
        const baseMarker = bullet ? "•" : decimalMarker;
        const rendering = {
          level: 0,
          numId: 1,
          isBullet: bullet,
          numFmt: kind,
          markerTemplate,
          levelStarts: [start],
          ...(bullet
            ? { markerFormatting: { fontFamily: { ascii: "Symbol", hAnsi: "Symbol" } } }
            : {}),
        };
        const numPr = paragraphNumberingReference({ numId: 1, ilvl: 0 });
        const doc = schema.node("doc", null, [
          schema.node(
            "paragraph",
            {
              numPr,
              ...listRenderingAttrPatch({
                ...rendering,
                marker:
                  templateKind === "cleared" ? `${baseMarker}\t${suffix}` : `outdated\t${suffix}`,
                foldedMarkerSuffix: suffix,
              }),
              ...(templateKind === "cleared" ? { listMarkerTemplate: null } : {}),
            },
            schema.text("first"),
          ),
          schema.node(
            "paragraph",
            {
              numPr,
              ...CLEARED_LIST_RENDERING_ATTRS,
              ...listRenderingAttrPatch({
                ...rendering,
                marker: templateKind === "literal" ? "Article" : "outdated",
              }),
            },
            schema.text("second"),
          ),
        ]);
        const markers = toFlowBlocks(doc)
          .filter((block) => block.kind === "paragraph")
          .map((block) => block.attrs?.listMarker);
        expect(markers).toEqual([`${baseMarker}\t${suffix}`, bullet ? "•" : nextDecimalMarker]);
      },
    ),
    { numRuns: 50 },
  );
});
