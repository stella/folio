import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { EditorState } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";

import {
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  type EditorIntent,
  OP_STORIES,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Document, ParagraphContent } from "../types/document";
import { createHarnessState, harnessRuntimeManager } from "./editorHarness";

setDefaultTimeout(propertyTestTimeout(120_000));

const PARA_ID = "12345678";
const HREF = "https://x.example/";

type Piece =
  | { kind: "run"; text: string }
  | { kind: "control" | "bidi" | "smartTag"; text: string }
  | { kind: "math" }
  | { kind: "move" };

const pieceArbitrary: fc.Arbitrary<Piece> = fc.oneof(
  fc.record({
    kind: fc.constant("run" as const),
    text: fc.constantFrom("ab", "c"),
  }),
  fc.record({
    kind: fc.constantFrom(
      "control" as const,
      "bidi" as const,
      "smartTag" as const,
    ),
    text: fc.constantFrom("de", "f"),
  }),
  fc.record({ kind: fc.constantFrom("math" as const, "move" as const) }),
);

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const documentOf = (pieces: readonly Piece[]): Document => {
  let id = 1;
  const content: ParagraphContent[] = [];
  for (const piece of pieces) {
    switch (piece.kind) {
      case "run":
        content.push(run(piece.text));
        break;
      case "control":
        content.push({
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 100 + id++ },
          content: [run(piece.text)],
        });
        break;
      case "bidi":
        content.push({
          type: "inlineWrapper",
          kind: "bidi",
          control: "embedding",
          direction: "rtl",
          content: [run(piece.text)],
        });
        break;
      case "smartTag":
        content.push({
          type: "inlineWrapper",
          kind: "smartTag",
          element: "place",
          content: [run(piece.text)],
        });
        break;
      case "math":
        content.push({
          type: "mathEquation",
          display: "inline",
          ommlXml:
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>',
          plainText: "x",
        });
        break;
      case "move": {
        const marker = id++;
        content.push(
          {
            type: "moveFromRangeStart",
            id: marker,
            name: `move${String(marker)}`,
            author: "Other",
            date: "2026-01-01T00:00:00Z",
          },
          { type: "moveFromRangeEnd", id: marker },
        );
        break;
      }
      default: {
        const unreachable: never = piece;
        return unreachable;
      }
    }
  }
  return {
    package: {
      document: {
        content: [{ type: "paragraph", paraId: PARA_ID, content }],
      },
    },
  };
};

/** Each unit's text or equation, the link over it, and the containers around it. */
type Unit = { unit: string; link: string | undefined; within: string };

const units = (document: Document): Unit[] => {
  const out: Unit[] = [];
  const walk = (
    items: readonly ParagraphContent[],
    link: string | undefined,
    within: string,
  ): void => {
    for (const item of items) {
      switch (item.type) {
        case "run":
          for (const part of item.content)
            if (part.type === "text")
              for (const unit of part.text) out.push({ unit, link, within });
          break;
        case "hyperlink":
          walk(item.children, item.href, within);
          break;
        case "inlineSdt":
          walk(
            item.content,
            link,
            `${within}/sdt${String(item.properties.id)}`,
          );
          break;
        case "inlineWrapper":
          walk(item.content, link, `${within}/${item.kind}`);
          break;
        case "mathEquation":
          out.push({ unit: "<math>", link, within });
          break;
        default:
          break;
      }
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") walk(block.content, undefined, "");
  return out;
};

/** The ids of the content controls and move markers, in order. */
const records = (document: Document): string[] => {
  const out: string[] = [];
  const walk = (items: readonly ParagraphContent[]): void => {
    for (const item of items) {
      if (item.type === "inlineSdt") {
        out.push(`sdt${String(item.properties.id)}`);
        walk(item.content);
      } else if (item.type === "inlineWrapper") {
        out.push(item.kind);
        walk(item.content);
      } else if (item.type === "hyperlink") walk(item.children);
      else if (
        item.type === "moveFromRangeStart" ||
        item.type === "moveFromRangeEnd"
      )
        out.push(`${item.type}${String(item.id)}`);
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") walk(block.content);
  return out;
};

/** The PM position before unit `index` (or after the last one). */
const legacyPosition = (state: EditorState, index: number): number => {
  let seen = 0;
  let found: number | undefined;
  let end = 0;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    let width = 0;
    if (node.isText) width = node.text?.length ?? 0;
    else if (node.type.name === "math") width = 1;
    if (width > 0 && index < seen + width) {
      found = position + (index - seen);
      return false;
    }
    seen += width;
    if (width > 0) end = position + node.nodeSize;
    return true;
  });
  return found ?? end;
};

const legacySetLink = (
  document: Document,
  from: number,
  to: number,
): Document | undefined => {
  let state = createHarnessState(document, "editing");
  state = state.apply(
    state.tr.setSelection(
      TextSelection.create(
        state.doc,
        legacyPosition(state, from),
        to === from
          ? legacyPosition(state, from)
          : legacyPosition(state, to - 1) + 1,
      ),
    ),
  );
  let applied = false;
  harnessRuntimeManager(state).requireCommand("setHyperlink")(HREF)(
    state,
    (tr) => {
      state = state.apply(tr);
      applied = true;
    },
  );
  return applied
    ? fromProseDoc(state.doc, document, {
        stylesheetSource: { type: "package" },
      })
    : undefined;
};

const canonicalEdit = (document: Document, intent: EditorIntent) => {
  const ids = createEditorIntentIdAllocator()(document, intent);
  return compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: ids.newIds },
  })
    .andThen(({ ops }) => applyDocumentOps(document, ops))
    .unwrap();
};

const position = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: PARA_ID,
  offset,
});

test("direct link edits over content controls, wrappers, equations and move markers match the legacy editor", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(pieceArbitrary, { minLength: 1, maxLength: 6 }),
      fc.nat(),
      fc.nat(),
      async (pieces, left, right) => {
        const document = documentOf(pieces);
        const length = units(document).length;
        const from = Math.min(left % (length + 1), right % (length + 1));
        const to = Math.max(left % (length + 1), right % (length + 1));
        if (from === to) return;
        const set = canonicalEdit(document, {
          type: "setHyperlink",
          from: position(from),
          to: position(to),
          href: HREF,
        });
        const linked = units(set.document);
        // Every unit keeps its text and its containers; the range is linked.
        expect(linked.map(({ unit, within }) => [unit, within])).toEqual(
          units(document).map(({ unit, within }) => [unit, within]),
        );
        for (const [index, unit] of linked.entries())
          expect(unit.link).toBe(
            index >= from && index < to && unit.unit !== "<math>"
              ? HREF
              : undefined,
          );
        expect(records(set.document)).toEqual(records(document));
        // The legacy editor links the same units in the same containers; it
        // drops an equation inside the range, which this edit keeps unlinked.
        const legacy = legacySetLink(document, from, to);
        if (legacy !== undefined)
          expect(units(legacy).filter(({ unit }) => unit !== "<math>")).toEqual(
            linked.filter(({ unit }) => unit !== "<math>"),
          );
        // Exact inverse, and the pieces read as one link to remove.
        expect(
          applyDocumentOps(set.document, set.inverse).unwrap().document,
        ).toStrictEqual(document);
        const removed = canonicalEdit(set.document, {
          type: "removeHyperlink",
          from: position(from),
          to: position(to),
        });
        expect(units(removed.document)).toEqual(units(document));
        expect(records(removed.document)).toEqual(records(document));
        // Saved and reopened, the edit reads the same.
        const reopened = await parseDocx(await createDocx(set.document));
        expect(units(reopened)).toEqual(linked);
        expect(records(reopened)).toEqual(records(set.document));
      },
    ),
    {
      numRuns: 80,
      id: "direct link edits over content controls, wrappers, equations and move markers match the legacy editor",
    },
  );
});
