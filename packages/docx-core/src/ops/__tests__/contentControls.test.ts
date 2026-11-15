import { describe, expect, test } from "bun:test";

import type { Document, SdtProperties } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { blockControlHolding } from "../contentControls";
import { isBlockScoped } from "../sequencing/scope";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type SetContentControlOp,
} from "../types";

const dropdown: SdtProperties = {
  sdtType: "dropdown",
  id: 101,
  listItems: [
    { displayText: "California", value: "ca" },
    { displayText: "New York", value: "ny" },
  ],
  dropdownLastValue: "ca",
};
const checkbox: SdtProperties = {
  sdtType: "checkbox",
  id: 102,
  checked: false,
};

const text = (value: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text: value }],
});

/** A block dropdown in a table cell, then a paragraph holding an inline checkbox. */
const controls = (second: SdtProperties = checkbox): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [
                    {
                      type: "blockSdt",
                      properties: dropdown,
                      content: [
                        {
                          type: "paragraph",
                          paraId: "0000A001",
                          content: [text("California")],
                        },
                      ],
                    },
                    { type: "paragraph", paraId: "0000A002", content: [] },
                  ],
                },
              ],
            },
          ],
        },
        {
          type: "paragraph",
          paraId: "0000A003",
          content: [
            text("Agreed "),
            {
              type: "inlineSdt",
              properties: second,
              content: [text("☐")],
            },
          ],
        },
      ],
    },
  },
});

const setDropdown = {
  type: DOCUMENT_OP_TYPES.SET_CONTENT_CONTROL,
  story: OP_STORIES.MAIN,
  controlId: 101,
  blockIds: ["0000A001"],
  expected: dropdown,
  properties: { ...dropdown, dropdownLastValue: "ny" },
} as const satisfies SetContentControlOp;
const setCheckbox = {
  type: DOCUMENT_OP_TYPES.SET_CONTENT_CONTROL,
  story: OP_STORIES.MAIN,
  controlId: 102,
  blockIds: ["0000A003"],
  expected: checkbox,
  properties: { ...checkbox, checked: true },
} as const satisfies SetContentControlOp;

const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

describe("setContentControl", () => {
  for (const [kind, op, read] of [
    [
      "a block-level control in a table cell",
      setDropdown,
      (document: Document) =>
        blockControlHolding(document.package.document.content, "0000A001")
          ?.control.properties,
    ],
    [
      "an inline control",
      setCheckbox,
      (document: Document) => {
        const paragraph = document.package.document.content.at(1);
        const control =
          paragraph?.type === "paragraph"
            ? paragraph.content.find((item) => item.type === "inlineSdt")
            : undefined;
        return control?.type === "inlineSdt" ? control.properties : undefined;
      },
    ],
  ] as const) {
    test(`sets ${kind} and its inverse restores it exactly`, () => {
      const original = controls();
      const forward = applied(original, op);
      expect(read(forward.document)).toEqual(op.properties);
      expect(forward.touched.modified).toEqual([...op.blockIds]);
      let restored = forward.document;
      for (const inverse of forward.inverse)
        restored = applied(restored, inverse).document;
      expect(restored).toEqual(original);
      expect(isBlockScoped(op)).toBe(true);
    });
  }

  test("refuses a stale, missing, ambiguous or retyping change", () => {
    const reason = (document: Document, op: DocumentOp) => {
      const result = applyDocumentOp(document, op);
      return result.isErr() ? result.error.reason : "applied";
    };
    expect(
      reason(controls(), {
        ...setDropdown,
        expected: { ...dropdown, dropdownLastValue: "ny" },
      }),
    ).toBe("stale");
    expect(reason(controls(), { ...setDropdown, blockIds: ["0000A002"] })).toBe(
      "stale",
    );
    expect(reason(controls(), { ...setDropdown, controlId: 999 })).toBe(
      "invalidOperation",
    );
    expect(
      reason(controls(), {
        ...setDropdown,
        controlId: 999,
        expected: { ...dropdown, id: 999 },
        properties: { ...dropdown, id: 999 },
      }),
    ).toBe("blockNotFound");
    expect(
      reason(controls({ ...checkbox, id: 101 }), {
        ...setCheckbox,
        controlId: 101,
        expected: { ...checkbox, id: 101 },
        properties: { ...checkbox, id: 101, checked: true },
      }),
    ).toBe("duplicateRecordId");
    expect(
      reason(controls(), {
        ...setCheckbox,
        properties: { ...checkbox, sdtType: "dropdown" },
      }),
    ).toBe("invalidOperation");
  });
});
