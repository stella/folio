/**
 * Lifting the wrapper must not lift the revision with it.
 *
 * The projection turns a `w:bdo`/`w:dir` tree into one inline sequence and
 * records the wrapper on the leaves it held. What it may not drop is the
 * revision the wrapper sat inside: text authored as
 * `<w:ins><w:bdo>x</w:bdo></w:ins>` is inserted text, and it has to reach the
 * editor carrying the insertion mark that says so. The stack the same leaf
 * carries is asserted in `toProseDoc-inline-wrapper.property.test.ts`.
 */

import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { toProseDoc } from "./toProseDoc";

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", paraId: "C0000001", content }],
      },
    },
  };
};

const markNamesOverText = (content: Paragraph["content"], text: string): string[] => {
  const names: string[] = [];
  toProseDoc(documentWith(content)).descendants((node) => {
    if (node.isText && node.text === text) {
      names.push(...node.marks.map((mark) => mark.type.name));
    }
  });
  return names;
};

const wrapperStackOverText = (content: Paragraph["content"], text: string): unknown => {
  let stack: unknown;
  toProseDoc(documentWith(content)).descendants((node) => {
    if (node.isText && node.text === text) {
      stack = node.marks.find((mark) => mark.type.name === "inlineWrapper")?.attrs["stack"];
    }
  });
  return stack;
};

const INFO = { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" };
const RUN = { type: "run", content: [{ type: "text", text: "x" }] } as const;

describe("a revision that holds a bidirectional wrapper", () => {
  test("the wrapped text reaches the editor marked inserted", () => {
    expect(
      markNamesOverText(
        [
          {
            type: "insertion",
            info: INFO,
            content: [
              {
                type: "inlineWrapper",
                kind: "bidi",
                control: "override",
                direction: "rtl",
                content: [RUN],
              },
            ],
          },
        ],
        "x",
      ),
    ).toContain("insertion");
  });

  test("the wrapped text reaches the editor marked deleted", () => {
    expect(
      markNamesOverText(
        [
          {
            type: "deletion",
            info: INFO,
            content: [
              {
                type: "inlineWrapper",
                kind: "bidi",
                control: "embedding",
                direction: "ltr",
                content: [RUN],
              },
            ],
          },
        ],
        "x",
      ),
    ).toContain("deletion");
  });

  test("the other authored order still marks its text", () => {
    expect(
      markNamesOverText(
        [
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "override",
            direction: "rtl",
            content: [{ type: "insertion", info: INFO, content: [RUN] }],
          },
        ],
        "x",
      ),
    ).toContain("insertion");
  });

  test("a wrapper on each side of the revision reaches the leaf, outermost first", () => {
    expect(
      wrapperStackOverText(
        [
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            direction: "rtl",
            content: [
              {
                type: "insertion",
                info: INFO,
                content: [
                  {
                    type: "inlineWrapper",
                    kind: "bidi",
                    control: "override",
                    direction: "ltr",
                    content: [RUN],
                  },
                ],
              },
            ],
          },
        ],
        "x",
      ),
    ).toEqual([
      { kind: "bidi", control: "embedding", direction: "rtl" },
      { kind: "bidi", control: "override", direction: "ltr" },
    ]);
  });

  test("a wrapper on each side of a content control reaches the leaf, outermost first", () => {
    expect(
      wrapperStackOverText(
        [
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            direction: "rtl",
            content: [
              {
                type: "inlineSdt",
                properties: { sdtType: "richText", tag: "bound" },
                content: [
                  {
                    type: "inlineWrapper",
                    kind: "bidi",
                    control: "override",
                    direction: "ltr",
                    content: [RUN],
                  },
                ],
              },
            ],
          },
        ],
        "x",
      ),
    ).toEqual([
      { kind: "bidi", control: "embedding", direction: "rtl" },
      { kind: "bidi", control: "override", direction: "ltr" },
    ]);
  });

  test("a content control inside a revision keeps its node", () => {
    const names: string[] = [];
    toProseDoc(
      documentWith([
        {
          type: "insertion",
          info: INFO,
          content: [
            {
              type: "inlineSdt",
              properties: { sdtType: "richText", tag: "bound" },
              content: [RUN],
            },
          ],
        },
      ]),
    ).descendants((node) => {
      names.push(node.type.name);
    });
    expect(names).toContain("sdt");
  });
});
