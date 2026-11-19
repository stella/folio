/**
 * A link built through the model API whose transparent wrapper holds content
 * a link cannot hold: an equation, a revision, a simple field or an inline
 * content control. The model types a wrapper's content as any paragraph
 * content, so nothing refuses the shape at construction; the editor reads it
 * as the parser reads a wrapper under a link's revision, with the wrapper
 * owning the link's pieces, and keeps every leaf.
 */

import { expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  DOCUMENT_OP_TYPES,
  inlineLeafSpans,
  OP_STORIES,
} from "@stll/docx-core/ops";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type {
  Document,
  Hyperlink,
  InlineWrapper,
  Paragraph,
  ParagraphContent,
  Run,
} from "../types/document";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const LINK = { type: "hyperlink", anchor: "target", href: "#target" } as const;
const link = (children: Hyperlink["children"]): Hyperlink => ({
  ...LINK,
  children,
});
const bidi = (content: ParagraphContent[]): InlineWrapper => ({
  type: "inlineWrapper",
  kind: "bidi",
  control: "embedding",
  content,
});

const HELD = {
  equation: () => ({
    type: "mathEquation",
    display: "inline",
    ommlXml: "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
    plainText: "x",
  }),
  insertion: (text: string) => ({
    type: "insertion",
    info: { id: 7, author: "Reviewer" },
    content: [run(text)],
  }),
  deletion: (text: string) => ({
    type: "deletion",
    info: { id: 8, author: "Reviewer" },
    content: [run(text)],
  }),
  field: (text: string) => ({
    type: "simpleField",
    instruction: "PAGE",
    fieldType: "PAGE",
    content: [run(text)],
  }),
  control: (text: string) => ({
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 9 },
    content: [run(text)],
  }),
} as const satisfies Record<string, (text: string) => ParagraphContent>;

const paragraphOf = (content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId: "00000001",
  content,
});
const documentOf = (paragraph: Paragraph): Document => ({
  package: { document: { content: [paragraph] } },
});

/** The paragraph's leaves in order: text as text, every other leaf by type. */
const leaves = (content: readonly ParagraphContent[]): string =>
  inlineLeafSpans(content)
    .map(({ node }) => (node.type === "text" ? node.text : `<${node.type}>`))
    .join("");

/** The text of every run a link holds, at any depth, in order. */
const linkedText = (
  items: readonly ParagraphContent[],
  linked = false,
): string =>
  items
    .map((item) => {
      switch (item.type) {
        case "run":
          return linked
            ? item.content
                .map((part) => (part.type === "text" ? part.text : ""))
                .join("")
            : "";
        case "hyperlink":
          return linkedText(item.children, true);
        case "inlineWrapper":
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
        case "inlineSdt":
        case "simpleField":
          return linkedText(item.content, linked);
        default:
          return "";
      }
    })
    .join("");

const throughTheEditor = (paragraph: Paragraph): Paragraph => {
  const source = documentOf(paragraph);
  const block = fromProseDoc(toProseDoc(source), source, {
    stylesheetSource: { type: "package" },
  }).package.document.content.at(0);
  if (block?.type !== "paragraph")
    throw new TypeError("The round trip lost its paragraph.");
  return block;
};

/** Whether an editor paragraph holds the node a held kind converts to. */
const holds = (paragraph: PMNode, kind: keyof typeof HELD): boolean => {
  let found = false;
  paragraph.descendants((node) => {
    const marks = node.marks.map(({ type }) => type.name);
    switch (kind) {
      case "equation":
        found ||= node.type.name === "math";
        break;
      case "insertion":
      case "deletion":
        found ||= node.text === "b" && marks.includes(kind);
        break;
      case "field":
        found ||= node.type.name === "structuredField";
        break;
      case "control":
        found ||= node.type.name === "sdt" && node.textContent === "b";
        break;
    }
  });
  return found;
};

const textArbitrary = fc.constantFrom("a", "b", "c");
const heldArbitrary: fc.Arbitrary<ParagraphContent> = fc
  .tuple(
    fc.constantFrom(...(Object.keys(HELD) as (keyof typeof HELD)[])),
    textArbitrary,
  )
  .map(([kind, text]) => HELD[kind](text));
const wrapperArbitrary: fc.Arbitrary<InlineWrapper> = fc
  .array(
    fc.oneof(
      textArbitrary.map(run),
      heldArbitrary,
      fc
        .array(fc.oneof(textArbitrary.map(run), heldArbitrary), {
          minLength: 1,
          maxLength: 3,
        })
        .map(bidi),
    ),
    { minLength: 1, maxLength: 4 },
  )
  .map(bidi);
const paragraphArbitrary: fc.Arbitrary<Paragraph> = fc
  .record({
    children: fc.array(fc.oneof(textArbitrary.map(run), wrapperArbitrary), {
      minLength: 1,
      maxLength: 4,
    }),
    revised: fc.boolean(),
  })
  .map(({ children, revised }) =>
    paragraphOf(
      revised
        ? [
            {
              type: "insertion",
              info: { id: 1, author: "Reviewer" },
              content: [link(children)],
            },
          ]
        : [link(children)],
    ),
  );

test(
  "the editor keeps every leaf of a link whose wrapper holds what a link cannot",
  () => {
    fc.assert(
      fc.property(paragraphArbitrary, (paragraph) => {
        const back = throughTheEditor(paragraph);
        expect(leaves(back.content)).toBe(leaves(paragraph.content));
        // Every run the link held is still linked.
        expect(linkedText(back.content)).toBe(linkedText(paragraph.content));
      }),
      propertyConfig({ numRuns: 200 }),
    );
  },
  propertyTestTimeout(30_000),
);

for (const kind of Object.keys(HELD) as (keyof typeof HELD)[]) {
  const paragraph = () =>
    paragraphOf([
      link([run("L"), bidi([run("a"), HELD[kind]("b")]), run("R")]),
    ]);

  test(`the editor keeps ${kind} in a wrapper of a built link`, () => {
    expect(holds(toProseDoc(documentOf(paragraph())).child(0), kind)).toBe(
      true,
    );
    const back = throughTheEditor(paragraph());
    expect(leaves(back.content)).toBe(leaves(paragraph().content));
  });

  test(`a canonical session opens on ${kind} in a wrapper of a built link`, () => {
    const session = createCanonicalSession(documentOf(paragraph()));
    expect(session.isOk() ? null : session.error.message).toBeNull();
  });

  test(`a canonical session shows ${kind} an operation inserts in a wrapper of a link`, () => {
    const session = createCanonicalSession(
      documentOf(paragraphOf([run("Body")])),
    ).unwrap();
    const state = EditorState.create({ schema, doc: session.projection.doc });
    const commit = session
      .prepareOperations(state, [
        {
          type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
          story: OP_STORIES.MAIN,
          at: { type: "after", blockId: "00000001" },
          blocks: [{ ...paragraph(), paraId: "00000002" }],
        },
      ])
      .unwrap();
    expect(holds(commit.projection.doc.child(1), kind)).toBe(true);
    expect(commit.publish().isOk()).toBe(true);
  });
}
