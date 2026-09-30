import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { rebaseParagraphRunContent, rebaseParagraphRuns } from "./rebaseParagraphRuns";
import { schema } from "./schema";

test(
  "an unchanged cascade preserves run carriers and their authored provenance",
  () => {
    assertProperty(
      fc.property(fc.boolean(), fc.boolean(), fc.boolean(), (bold, rtl, explicit) => {
        const marks = [
          ...(bold ? [schema.mark("bold")] : []),
          ...(rtl ? [schema.mark("rtl")] : []),
          ...(explicit
            ? [schema.mark("runFormattingOverride", { _authoredOn: ["bold", "rtl"] })]
            : []),
        ];
        const paragraph = schema.nodes["paragraph"]!.create({ paraId: "source" }, [
          schema.text("Text", marks),
          schema.nodes["tab"]!.create(null, null, marks),
        ]);
        const target = schema.nodes["paragraph"]!.create({ paraId: "target" });
        let rebases = 0;
        const content = rebaseParagraphRunContent({
          paragraph,
          target,
          position: 0,
          styleResolver: null,
          onRebased: () => {
            rebases += 1;
          },
        });
        expect(content).toBe(paragraph.content);
        expect(rebases).toBe(0);
        const state = EditorState.create({ doc: schema.node("doc", null, paragraph) });
        const tr = rebaseParagraphRuns({
          tr: state.tr,
          position: 0,
          previous: paragraph,
          target,
          styleResolver: null,
        });
        expect(tr.steps).toHaveLength(0);
        expect(tr.doc).toBe(state.doc);
      }),
      { numRuns: 32 },
    );
  },
  propertyTestTimeout(),
);
