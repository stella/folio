/**
 * A content control's selection state survives a rebuild without gaining one.
 *
 * `w:dropDownList@w:lastValue` and `w:comboBox@w:lastValue` default to the
 * empty string in the schema, so "never selected", "cleared" and "selected"
 * are three distinguishable states. folio used to recover a missing selection
 * from the control's body text whenever no captured `w:sdtPr` was there to say
 * otherwise — which is every control an editor command built, and every
 * control on the rebuild path — so a dropdown showing its placeholder came
 * back saved as a selection nobody made (`dropdownLastValue: absent became
 * "Drop down1"` in the corpus census).
 *
 * The property covers every SDT kind crossed with every selection state, over
 * both paths that make the real serializer run:
 *
 *   rebuild  parse → drop the captured `w:sdtPr` → save → parse.
 *   editor   parse → toProseDoc → fromProseDoc → save → parse.
 *
 * `w:date@w:fullDate` and the checkbox state ride along: none of the three may
 * be invented, and none may be dropped.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, BlockSdt, Document, SdtProperties } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const SDT_KINDS = [
  "dropdown",
  "comboBox",
  "date",
  "checkbox",
  "plainText",
  "richText",
] as const satisfies readonly SdtProperties["sdtType"][];

/** Absent, the schema default spelled out, and a real pick. */
const SELECTIONS = ["absent", "empty", "value"] as const;
type Selection = (typeof SELECTIONS)[number];

const LIST_ITEMS = [
  { displayText: "Drop down1", value: "first" },
  { displayText: "Drop down1", value: "second" },
] as const;

/** The three facts a selection is made of, as the parser reports them. */
type SelectionState = {
  dropdownLastValue?: string;
  dateValueISO?: string;
  checked?: boolean;
};

type SdtCase = {
  kind: (typeof SDT_KINDS)[number];
  selection: Selection;
};

const selectionValue = (selection: Selection, picked: string): string | undefined => {
  switch (selection) {
    case "absent":
      return undefined;
    case "empty":
      return "";
    case "value":
      return picked;
    default:
      return selection satisfies never;
  }
};

/**
 * `w:fullDate` is an `xsd:dateTime`, so it has no empty form: a date control
 * with no pick writes no attribute. `w14:checked` is a boolean, so it has
 * none either. Only `@w:lastValue` carries all three states.
 */
const propertiesFor = ({ kind, selection }: SdtCase): SdtProperties => {
  switch (kind) {
    case "dropdown":
    case "comboBox": {
      const lastValue = selectionValue(selection, "second");
      return {
        sdtType: kind,
        listItems: [...LIST_ITEMS],
        ...(lastValue !== undefined ? { dropdownLastValue: lastValue } : {}),
      };
    }
    case "date":
      return {
        sdtType: "date",
        dateFormat: "d MMMM yyyy",
        ...(selection === "value" ? { dateValueISO: "2026-06-02T00:00:00Z" } : {}),
      };
    case "checkbox":
      return { sdtType: "checkbox", checked: selection === "value" };
    default:
      return { sdtType: kind };
  }
};

const selectionOf = ({
  dropdownLastValue,
  dateValueISO,
  checked,
}: SdtProperties): SelectionState => ({
  ...(dropdownLastValue !== undefined ? { dropdownLastValue } : {}),
  ...(dateValueISO !== undefined ? { dateValueISO } : {}),
  ...(checked !== undefined ? { checked } : {}),
});

/** Body text that looks exactly like a selection, so inference would show. */
const sdtBlock = (properties: SdtProperties): BlockSdt => ({
  type: "blockSdt",
  properties,
  content: [
    {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "Drop down1" }] }],
    },
  ],
});

const sdtCase = fc.record({
  kind: fc.constantFrom(...SDT_KINDS),
  selection: fc.constantFrom(...SELECTIONS),
});

const withSdt = (document: Document, block: BlockContent): Document => ({
  ...document,
  package: {
    ...document.package,
    document: { ...document.package.document, content: [block] },
  },
});

const readSelection = (document: Document): SelectionState | null => {
  const block = document.package.document.content.at(0);
  if (!block || block.type !== "blockSdt") {
    return null;
  }
  return selectionOf(block.properties);
};

/** The capture the `reserialize` invariant strips, so the serializer runs. */
const withoutCapturedProperties = (document: Document): Document => {
  const block = document.package.document.content.at(0);
  if (!block || block.type !== "blockSdt") {
    return document;
  }
  const { rawPropertiesXml: _raw, rawEndPropertiesXml: _rawEnd, ...properties } = block.properties;
  return withSdt(document, { ...block, properties });
};

const parse = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { detectVariables: false, preloadFonts: false });

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

describe("content control selection state survives a rebuild", () => {
  test("a dropdown nobody touched is not saved with a selection", async () => {
    const template = await parse(await createEmptyDocx());
    const authored = withSdt(
      template,
      sdtBlock({ sdtType: "dropdown", listItems: [...LIST_ITEMS] }),
    );

    const opened = await parse(await save(authored));
    expect(readSelection(opened)).toEqual({});

    const rebuilt = await parse(await save(withoutCapturedProperties(opened)));
    expect(readSelection(rebuilt)).toEqual({});
  });

  test(
    "every kind × selection state is preserved exactly, on both rebuild paths",
    async () => {
      const template = await parse(await createEmptyDocx());

      await fc.assert(
        fc.asyncProperty(sdtCase, async (sdt) => {
          // The authored model is the ground truth: reading it back from the
          // first save would bake an invented selection into the expectation.
          const properties = propertiesFor(sdt);
          const expected = selectionOf(properties);

          const opened = await parse(await save(withSdt(template, sdtBlock(properties))));
          expect(readSelection(opened)).toEqual(expected);

          const rebuilt = await parse(await save(withoutCapturedProperties(opened)));
          expect(readSelection(rebuilt)).toEqual(expected);

          const edited = fromProseDoc(toProseDoc(opened), opened);
          const afterEditor = await parse(await save(edited));
          expect(readSelection(afterEditor)).toEqual(expected);
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(60_000),
  );
});
