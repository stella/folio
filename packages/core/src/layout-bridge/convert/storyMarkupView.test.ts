import { panic } from "better-result";
import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import type { OpStory } from "@stll/docx-core/ops";

import { displayCommentsFrom } from "../../display-list/build/commentAnnotations";
import type { FlowBlock } from "../../layout-engine/types";
import type { DisplayMode } from "../../managers/EditorModeManager";
import { proseDocToBlocks } from "../../prosemirror/conversion/fromProseDoc";
import { projectMarkupView } from "../../prosemirror/markupViewProjection";
import { schema } from "../../prosemirror/schema";
import { convertNoteStoryToFlowBlocks } from "./footnoteLayout";
import { convertHeaderFooterPmDocToContent } from "./headerFooterLayout";
import { toFlowBlocks } from "./toFlowBlocks";

type IndependentStoryKind =
  Extract<OpStory, string> | Exclude<OpStory, string>["kind"];
type RenderStory = (doc: PMNode, markupView: DisplayMode) => FlowBlock[];
const renderHeaderFooter: RenderStory = (doc, markupView) => {
  const result = convertHeaderFooterPmDocToContent(
    doc,
    400,
    {
      section: "header",
      pageSize: { w: 600, h: 800 },
      margins: {
        top: 50,
        bottom: 50,
        left: 50,
        right: 50,
        header: 20,
        footer: 20,
      },
    },
    { markupView, measureBlocks: () => [] },
  );
  if (!result) panic("Expected header/footer content");
  return result.blocks;
};
const renderNote: RenderStory = (doc, markupView) =>
  convertNoteStoryToFlowBlocks(proseDocToBlocks(doc, []), { markupView })
    .flowBlocks;

// Independent stories come from the operation contract. Text boxes are nested
// content; comments are annotations, with no editable PM story (see OpStory).
const renderStories = {
  main: (doc, markupView) => {
    const projection = projectMarkupView(
      EditorState.create({ doc }),
      markupView,
    );
    return toFlowBlocks(projection.type === "resolved" ? projection.doc : doc);
  },
  header: renderHeaderFooter,
  footer: renderHeaderFooter,
  footnote: renderNote,
  endnote: renderNote,
} satisfies Record<IndependentStoryKind, RenderStory>;

const textOf = (blocks: FlowBlock[]): string =>
  blocks
    .map((block): string => {
      switch (block.kind) {
        case "paragraph":
          return block.runs
            .map((run) => (run.kind === "text" ? run.text : ""))
            .join("");
        case "textBox":
          return textOf(block.content);
        case "table":
          return block.rows
            .map((row) => row.cells.map((cell) => textOf(cell.blocks)).join(""))
            .join("");
        default:
          return "";
      }
    })
    .join("");

for (const [kind, render] of Object.entries(renderStories)) {
  for (const nested of [false, true]) {
    test(`${kind} projects ${nested ? "nested text box" : "story"} revisions`, () => {
      const paragraph = schema.node("paragraph", null, [
        schema.text("base"),
        schema.text("inserted", [schema.mark("insertion", { revisionId: 1 })]),
        schema.text("deleted", [schema.mark("deletion", { revisionId: 2 })]),
      ]);
      const doc = schema.node("doc", null, [
        nested ? schema.node("textBox", null, [paragraph]) : paragraph,
      ]);
      const authored = doc.toJSON();
      for (const markupView of [
        "no-markup",
        "original",
        "simple-markup",
        "all-markup",
      ] as const) {
        const expectedText = {
          original: "basedeleted",
          "all-markup": "baseinserteddeleted",
          "no-markup": "baseinserted",
          "simple-markup": "baseinserted",
        } as const satisfies Record<DisplayMode, string>;
        const expected = expectedText[markupView];
        expect(textOf(render(doc, markupView))).toBe(expected);
        expect(doc.toJSON()).toEqual(authored);
      }
    });
  }
}

test("comment summaries retain annotation text and omit preserved revision XML", () => {
  expect(
    displayCommentsFrom([
      {
        id: 1,
        author: "Reviewer",
        content: [
          {
            type: "paragraph",
            content: [
              { type: "run", content: [{ type: "text", text: "annotation" }] },
            ],
          },
        ],
        preserved: {
          children: [
            { index: 0, xml: "<w:ins><w:r><w:t>preserved</w:t></w:r></w:ins>" },
          ],
        },
      },
    ]),
  ).toEqual([{ id: 1, author: "Reviewer", contents: "annotation" }]);
});

// The body has a live state; detached stories must enter the same projection
// owner before flow conversion. Guard the production wiring as well as text.
const hasProjectionWiring = (source: string) =>
  /const projection = project(?:Story)?MarkupView\(/.test(source) &&
  /const viewDoc = projection(?:\?)?\.type === "resolved" \? projection.doc : (?:pmDoc|state.doc)/.test(
    source,
  ) &&
  /toFlowBlocks\((?:viewDoc|flowDoc),/.test(source);

test("every independent story renderer projects before flow conversion", async () => {
  const owners = {
    main: "../../controller/layoutPipeline.ts",
    header: "./headerFooterLayout.ts",
    footer: "./headerFooterLayout.ts",
    footnote: "./footnoteLayout.ts",
    endnote: "./footnoteLayout.ts",
  } satisfies Record<IndependentStoryKind, string>;
  expect(Object.keys(owners).sort()).toEqual(Object.keys(renderStories).sort());
  for (const owner of new Set(Object.values(owners))) {
    const source = await Bun.file(new URL(owner, import.meta.url)).text();
    expect(hasProjectionWiring(source)).toBe(true);
    expect(
      hasProjectionWiring(
        source.replace(/project(?:Story)?MarkupView\(/g, "unprojected("),
      ),
    ).toBe(false);
  }
});

// Read the registered story union, so adding a kind fails even though test
// sources are excluded from published-package typechecking.
test("the rendering census covers the registered story contract", async () => {
  const source = await Bun.file(
    new URL("./types.ts", import.meta.resolve("@stll/docx-core/ops")),
  ).text();
  const union = /export type OpStory = ([^;]+);/.exec(source)?.at(1);
  if (!union) panic("Expected the registered story union");
  const kinds: string[] = [];
  for (const member of union.split(" | ")) {
    if (member === "typeof OP_STORIES.MAIN") {
      const main = /MAIN: "([^"]+)"/.exec(source)?.at(1);
      if (!main) panic("Expected main story identity");
      kinds.push(main);
      continue;
    }
    const definition = new RegExp(`export type ${member} = { kind: ([^;]+);`)
      .exec(source)
      ?.at(1);
    if (!definition) panic("A story member needs a rendering census decision");
    for (const literal of definition.matchAll(/"([^"]+)"/g)) {
      const kind = literal.at(1);
      if (kind) kinds.push(kind);
    }
  }
  const declared = kinds.sort();
  const exercised = Object.keys(renderStories).sort();
  expect(exercised).toEqual(declared);
  expect(exercised.slice(1)).not.toEqual(declared);
});
