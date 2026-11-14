// Measurement only: what the legacy editor does with a direct link edit over
// content the canonical link edit refuses. Never merged.
import { setDefaultTimeout, test } from "bun:test";
import type { EditorState } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";

import type { Document, ParagraphContent } from "../types/document";
import {
  createHarnessState,
  harnessRuntimeManager,
  parseShapeDocument,
  saveHarnessState,
} from "./editorHarness";

setDefaultTimeout(600_000);

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const sdt = (content: ParagraphContent[]): ParagraphContent =>
  ({
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 101, alias: "Box" },
    content,
  }) as ParagraphContent;
const bidi = (content: ParagraphContent[]): ParagraphContent =>
  ({
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    direction: "rtl",
    content,
  }) as ParagraphContent;
const smartTag = (content: ParagraphContent[]): ParagraphContent =>
  ({
    type: "inlineWrapper",
    kind: "smartTag",
    element: "place",
    content,
  }) as ParagraphContent;
const math: ParagraphContent = {
  type: "mathEquation",
  display: "inline",
  ommlXml:
    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>',
  plainText: "x",
} as ParagraphContent;
const moveStart = {
  type: "moveFromRangeStart",
  id: 7,
  name: "move1",
  author: "A",
  date: "2026-01-01T00:00:00Z",
} as ParagraphContent;
const moveEnd = { type: "moveFromRangeEnd", id: 7 } as ParagraphContent;

const documentOf = (content: ParagraphContent[]): Document =>
  ({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "12345678", content }],
      },
    },
  }) as Document;

// Each shape: content, and the text offsets [from, to] selected.
const SHAPES: Record<string, [ParagraphContent[], number, number]> = {
  sdtCut: [[run("ab"), sdt([run("cd")]), run("ef")], 1, 3],
  sdtInside: [[run("ab"), sdt([run("cd")]), run("ef")], 2, 4],
  sdtAcross: [[run("ab"), sdt([run("cd")]), run("ef")], 1, 5],
  sdtElsewhere: [[run("ab"), sdt([run("cd")]), run("ef")], 4, 6],
  bidiCut: [[run("ab"), bidi([run("cd")]), run("ef")], 1, 3],
  bidiAcross: [[run("ab"), bidi([run("cd")]), run("ef")], 1, 5],
  smartTagCut: [[run("ab"), smartTag([run("cd")]), run("ef")], 1, 3],
  mathBetween: [[run("ab"), math, run("cd")], 1, 4],
  mathElsewhere: [[run("ab"), math, run("cd")], 0, 1],
  moveMarkers: [[run("ab"), moveStart, run("cd"), moveEnd, run("ef")], 1, 5],
};

/** The PM position of a text offset in the first textblock. */
const positionAt = (state: EditorState, offset: number): number => {
  let seen = 0;
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (!node.isText) return true;
    const length = node.text?.length ?? 0;
    if (offset <= seen + length) {
      found = position + (offset - seen);
      return false;
    }
    seen += length;
    return true;
  });
  return found ?? -1;
};

const shapeOf = (items: readonly unknown[]): unknown[] =>
  items.map((item) => {
    const record = item as Record<string, unknown>;
    const out: Record<string, unknown> = { t: record["type"] };
    if (record["type"] === "run")
      out["x"] = (record["content"] as { text?: string }[])
        .map((part) => part.text ?? "")
        .join("");
    if (record["href"] !== undefined) out["href"] = record["href"];
    if (record["kind"] !== undefined) out["kind"] = record["kind"];
    if (record["properties"] !== undefined)
      out["id"] = (record["properties"] as { id?: number }).id;
    if (Array.isArray(record["content"]) && record["type"] !== "run")
      out["c"] = shapeOf(record["content"]);
    if (Array.isArray(record["children"]))
      out["c"] = shapeOf(record["children"]);
    if (record["type"] === "mathEquation") out["x"] = record["plainText"];
    return out;
  });

test("S20LEGACYCUT legacy direct link edits over refused content", async () => {
  const report: Record<string, unknown> = {};
  for (const [name, [content, from, to]] of Object.entries(SHAPES)) {
    try {
      const document = documentOf(content);
      let state = createHarnessState(document, "editing");
      const before = state.doc.textContent;
      state = state.apply(
        state.tr.setSelection(
          TextSelection.create(
            state.doc,
            positionAt(state, from),
            positionAt(state, to),
          ),
        ),
      );
      const command =
        harnessRuntimeManager(state).requireCommand("setHyperlink")(
          "https://x.example/",
        );
      let applied = false;
      command(state, (tr) => {
        state = state.apply(tr);
        applied = true;
      });
      if (!applied) {
        report[name] = "refuses";
        continue;
      }
      const saved = await saveHarnessState(state, document);
      const reopened = await parseShapeDocument(saved.bytes);
      const paragraph = reopened.package.document.content.at(0);
      report[name] = {
        textKept: state.doc.textContent === before,
        reopened:
          paragraph?.type === "paragraph" ? shapeOf(paragraph.content) : null,
      };
    } catch (error) {
      report[name] = `throws ${String(error).slice(0, 200)}`;
    }
  }
  console.log(`S20LEGACYCUT ${JSON.stringify(report)}`);
});
