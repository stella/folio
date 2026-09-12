import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { expectHyperlinkMarkAttrs } from "../../prosemirror/attrs";
import { schema } from "../../prosemirror/schema";
import { applyExactInlineOwnership } from "./docx-text-executor";

describe("exact DOCX inline ownership", () => {
  test("does not collapse an inherited mark into a distinct adjacent target occurrence", () => {
    const attrs = {
      href: "https://example.invalid/same",
      tooltip: "same",
      target: "_self",
      history: false,
      docLocation: "clause-one",
      rId: "rId2",
      _docxHyperlinkIndex: 0,
    } as const;
    const inherited = schema.marks.hyperlink.create(attrs);
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("newold", [inherited])]),
      ]),
    });
    const tr = state.tr;

    applyExactInlineOwnership({
      tr,
      from: 1,
      to: 4,
      containers: [
        {
          type: "hyperlink",
          href: attrs.href,
          tooltip: attrs.tooltip,
          target: attrs.target,
          history: attrs.history,
          docLocation: attrs.docLocation,
          occurrence: { blockId: "target-block", index: 1 },
        },
      ],
    });

    const occurrences: ReturnType<typeof expectHyperlinkMarkAttrs>[] = [];
    tr.doc.descendants((node) => {
      if (!node.isText) return;
      const hyperlink = node.marks.find(({ type }) => type.name === "hyperlink");
      if (hyperlink) occurrences.push(expectHyperlinkMarkAttrs(hyperlink));
    });
    expect(occurrences.map(({ _docxHyperlinkIndex }) => _docxHyperlinkIndex)).toEqual([1, 0]);
    expect(occurrences.at(0)).toMatchObject({
      href: attrs.href,
      tooltip: attrs.tooltip,
      target: attrs.target,
      history: attrs.history,
      docLocation: attrs.docLocation,
    });
    expect(occurrences.at(0)?.rId).toBeUndefined();
    expect(occurrences.at(1)?.rId).toBe("rId2");
  });
});
