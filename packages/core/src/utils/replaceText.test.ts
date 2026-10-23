import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { FindReplaceManager } from "../managers/FindReplaceManager";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../types/document";
import { createEmptyDocument } from "./createDocument";
import {
  createDefaultFindOptions,
  findInDocument,
  getFindParagraphText,
} from "./findReplace";
import { replaceTextInDocument } from "./replaceText";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraphOf = (document: Document): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Fixture paragraph missing");
  return paragraph;
};

const prefixes = {
  simpleField: {
    type: "simpleField",
    fieldType: "DOCVARIABLE",
    instruction: "DOCVARIABLE value",
    content: [run("cached")],
    fldLock: true,
  },
  complexField: {
    type: "complexField",
    fieldType: "DOCVARIABLE",
    instruction: "DOCVARIABLE value",
    fieldCode: [
      {
        type: "run",
        content: [{ type: "instrText", text: "DOCVARIABLE value" }],
      },
    ],
    fieldResult: [run("cached")],
    fldLock: true,
  },
  nativeTokens: {
    type: "run",
    content: [
      { type: "text", text: "A" },
      { type: "tab" },
      { type: "symbol", font: "Arial", char: "00A7" },
      { type: "softHyphen" },
      { type: "noBreakHyphen" },
      { type: "break", breakType: "column" },
    ],
  },
  opaqueText: {
    type: "preservedInline",
    xml: "<w:customXml><w:r><w:t>reader</w:t></w:r></w:customXml>",
    text: "reader",
  },
} as const satisfies Record<string, ParagraphContent>;

describe("replaceTextInDocument authored offsets", () => {
  test.each(Object.entries(prefixes))(
    "replaces ordinary text after %s without changing source units",
    (_, prefix) => {
      const source: Document = {
        package: {
          document: {
            content: [
              { type: "paragraph", content: [prefix, run(" target tail")] },
            ],
          },
        },
      };
      const before = structuredClone(source);
      const beforePmPrefix = toProseDoc(source).child(0).child(0).toJSON();
      const match = findInDocument(
        source,
        "target",
        createDefaultFindOptions(),
      ).at(0);
      if (!match) panic("Fixture match missing");
      const result = replaceTextInDocument(
        source,
        {
          start: {
            paragraphIndex: match.paragraphIndex,
            offset: match.startOffset,
          },
          end: {
            paragraphIndex: match.paragraphIndex,
            offset: match.endOffset,
          },
        },
        "value",
      );
      expect(getFindParagraphText(paragraphOf(result))).toBe(
        getFindParagraphText(paragraphOf(source)).replace("target", "value"),
      );
      expect(paragraphOf(result).content.at(0)).toEqual(prefix);
      expect(toProseDoc(result).child(0).child(0).toJSON()).toEqual(
        beforePmPrefix,
      );
      expect(
        findInDocument(result, "target", createDefaultFindOptions()),
      ).toHaveLength(0);
      expect(source).toEqual(before);
    },
  );

  const containers = {
    hyperlink: {
      type: "hyperlink",
      href: "https://example.test",
      children: [run("target inside")],
    },
    transparent: {
      type: "inlineWrapper",
      kind: "smartTag",
      element: "tag",
      content: [run("target inside")],
    },
    insertion: {
      type: "insertion",
      info: { id: 1, author: "Author" },
      content: [run("target inside")],
    },
    deletion: {
      type: "deletion",
      info: { id: 2, author: "Author" },
      content: [run("target inside")],
    },
    moveFrom: {
      type: "moveFrom",
      info: { id: 3, author: "Author" },
      content: [run("target inside")],
    },
    moveTo: {
      type: "moveTo",
      info: { id: 4, author: "Author" },
      content: [run("target inside")],
    },
    control: {
      type: "inlineSdt",
      properties: {},
      content: [run("target inside")],
    },
  } as const satisfies Record<string, ParagraphContent>;

  test.each(Object.entries(containers))(
    "keeps replacement inside its %s carrier",
    (_, container) => {
      const source: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [run("outside "), container, run(" tail")],
              },
            ],
          },
        },
      };
      const before = structuredClone(source);
      const match = findInDocument(
        source,
        "target",
        createDefaultFindOptions(),
      ).at(0);
      if (!match) panic("Fixture carrier match missing");
      const result = replaceTextInDocument(
        source,
        {
          start: {
            paragraphIndex: match.paragraphIndex,
            offset: match.startOffset,
          },
          end: {
            paragraphIndex: match.paragraphIndex,
            offset: match.endOffset,
          },
        },
        "value",
      );
      expect(getFindParagraphText(paragraphOf(result))).toBe(
        "outside value inside tail",
      );
      const changed = paragraphOf(result).content.at(1);
      expect(changed?.type).toBe(container.type);
      if (!changed) panic("Replacement carrier missing");
      expect(
        getFindParagraphText({ type: "paragraph", content: [changed] }),
      ).toBe("value inside");
      expect(paragraphOf(result).content.at(0)).toEqual(run("outside "));
      expect(source).toEqual(before);
    },
  );

  test("replaces across adjacent text runs while preserving their formatting", () => {
    const source: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                { ...run("pre tar"), formatting: { bold: true } },
                { ...run("get tail"), formatting: { italic: true } },
              ],
            },
          ],
        },
      },
    };
    const result = replaceTextInDocument(
      source,
      {
        start: { paragraphIndex: 0, offset: 4 },
        end: { paragraphIndex: 0, offset: 10 },
      },
      "value",
    );
    expect(getFindParagraphText(paragraphOf(result))).toBe("pre value tail");
    expect(paragraphOf(result).content.at(0)).toMatchObject({
      formatting: { bold: true },
    });
    expect(paragraphOf(result).content.at(1)).toMatchObject({
      formatting: { italic: true },
    });
  });

  test("explicit replacement formatting keeps suffix text before trailing native tokens", () => {
    const source: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  formatting: { italic: true },
                  content: [
                    { type: "text", text: "pre target suffix" },
                    { type: "tab" },
                    { type: "text", text: "tail" },
                  ],
                },
              ],
            },
          ],
        },
      },
    };
    const result = replaceTextInDocument(
      source,
      {
        start: { paragraphIndex: 0, offset: 4 },
        end: { paragraphIndex: 0, offset: 10 },
      },
      "value",
      { bold: true },
    );
    expect(getFindParagraphText(paragraphOf(result))).toBe(
      "pre value suffix\ttail",
    );
    expect(paragraphOf(result).content.at(1)).toMatchObject({
      formatting: { bold: true },
    });
    expect(paragraphOf(result).content.at(2)).toMatchObject({
      formatting: { italic: true },
    });
  });

  test.each([
    { type: "tab" },
    { type: "symbol", font: "Arial", char: "00A7" },
    { type: "softHyphen" },
    { type: "noBreakHyphen" },
  ] as const satisfies readonly Run["content"][number][])(
    "replaces selected $type native units using authored offsets",
    (unit) => {
      const source: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  {
                    type: "run",
                    content: [
                      { type: "text", text: "A" },
                      unit,
                      { type: "text", text: "B" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };
      const result = replaceTextInDocument(
        source,
        {
          start: { paragraphIndex: 0, offset: 1 },
          end: { paragraphIndex: 0, offset: 2 },
        },
        "value",
      );
      expect(getFindParagraphText(paragraphOf(result))).toBe("AvalueB");
      expect(paragraphOf(source).content.at(0)).toMatchObject({
        content: [
          { type: "text", text: "A" },
          unit,
          { type: "text", text: "B" },
        ],
      });
    },
  );

  test.each([
    {
      type: "preservedInline",
      xml: "<w:customXml><w:r><w:t>opaque</w:t></w:r></w:customXml>",
      text: "opaque",
    },
    {
      type: "run",
      content: [
        {
          type: "preservedXml",
          xml: "<w:ruby><w:rubyBase><w:r><w:t>opaque</w:t></w:r></w:rubyBase></w:ruby>",
          text: "opaque",
        },
      ],
    },
  ] as const satisfies readonly ParagraphContent[])(
    "opaque text substrings remain blocked with source metadata intact",
    (content) => {
      const source: Document = {
        package: {
          document: { content: [{ type: "paragraph", content: [content] }] },
        },
      };
      const before = structuredClone(source);
      const beforePm = toProseDoc(source).toJSON();
      const matches = findInDocument(source, "paq", createDefaultFindOptions());
      expect(matches).toHaveLength(1);
      expect(matches.at(0)?.replacement).toEqual({
        status: "blocked",
        reason: "atomic-content",
      });
      const manager = new FindReplaceManager<(typeof matches)[number]>();
      manager.setMatches(matches);
      expect(manager.replaceCurrent(source, "value")).toBeNull();
      expect(manager.replaceAll(source, matches, "value")).toEqual({
        document: source,
        replacedCount: 0,
        skippedCount: 1,
      });
      expect(source).toEqual(before);
      expect(toProseDoc(source).toJSON()).toEqual(beforePm);
    },
  );
});

const documentWithRevision = () => {
  const template = createEmptyDocument();
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "C0000001",
    formatting: { alignment: "center" },
    content: [
      run("left"),
      {
        type: "deletion",
        info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
        content: [run(" hidden ")],
      },
      run("right"),
    ],
  };
  return {
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content: [paragraph] },
    },
  };
};

type RevisionKind = Extract<ParagraphContent, { info: unknown }>["type"];

const revisionWrapperFactories = {
  insertion: (text: string): ParagraphContent => ({
    type: "insertion",
    info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
    content: [run(text)],
  }),
  deletion: (text: string): ParagraphContent => ({
    type: "deletion",
    info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
    content: [run(text)],
  }),
  moveFrom: (text: string): ParagraphContent => ({
    type: "moveFrom",
    info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
    content: [run(text)],
  }),
  moveTo: (text: string): ParagraphContent => ({
    type: "moveTo",
    info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
    content: [run(text)],
  }),
} as const satisfies Record<RevisionKind, (text: string) => ParagraphContent>;

const revisionCases = Object.entries(revisionWrapperFactories).flatMap(
  ([kind, wrap]) =>
    ["stock", "clause", "review"].map((needle) => ({ kind, wrap, needle })),
);

const documentWithContent = (content: ParagraphContent[]) => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph" as const, paraId: "C0000001", content }],
      },
    },
  };
};

describe("replaceTextInDocument", () => {
  test("resolves replacement offsets through tracked revision wrappers", () => {
    const source = documentWithRevision();
    const originalContent = structuredClone(source.package.document.content);
    const replaced = replaceTextInDocument(
      source,
      {
        start: { paragraphIndex: 0, offset: 0 },
        end: { paragraphIndex: 0, offset: 4 },
      },
      "replacement",
    );
    const paragraph = replaced.package.document.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") {
      panic("Expected the replaced paragraph");
    }
    expect(toProseDoc(replaced).textContent).toBe("replacement hidden right");
    expect(paragraph.content.some((item) => item.type === "deletion")).toBe(
      true,
    );
    expect(source.package.document.content).toEqual(originalContent);
    expect(paragraph.formatting).toEqual({ alignment: "center" });
  });

  test.each(revisionCases)(
    "replaces text inside $kind revision",
    ({ kind, wrap, needle }) => {
      const original = documentWithContent([
        run("prefix "),
        wrap(`before ${needle} after`),
        run(" suffix"),
      ]);
      const originalContent = structuredClone(
        original.package.document.content,
      );
      const sourceText = toProseDoc(original).textContent;
      const match = findInDocument(
        original,
        needle,
        createDefaultFindOptions(),
      ).at(0);
      if (match === undefined) {
        throw new Error(`Expected to find ${needle} in ${kind}`);
      }
      const replaced = replaceTextInDocument(
        original,
        {
          start: {
            paragraphIndex: match.paragraphIndex,
            offset: match.startOffset,
          },
          end: {
            paragraphIndex: match.paragraphIndex,
            offset: match.endOffset,
          },
        },
        "shares",
      );

      expect(toProseDoc(original).textContent).toBe(sourceText);
      expect(original.package.document.content).toEqual(originalContent);
      expect(toProseDoc(replaced).textContent).toBe(
        "prefix before shares after suffix",
      );
      const paragraph = replaced.package.document.content.at(0);
      expect(paragraph?.type).toBe("paragraph");
      if (paragraph?.type !== "paragraph") {
        panic("Expected the replaced paragraph");
      }
      expect(paragraph.content.at(1)?.type).toBe(kind);
    },
  );

  test("applies requested formatting to replacement text", () => {
    const replaced = replaceTextInDocument(
      documentWithRevision(),
      {
        start: { paragraphIndex: 0, offset: 0 },
        end: { paragraphIndex: 0, offset: 4 },
      },
      "replacement",
      { bold: true },
    );
    const paragraph = replaced.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") {
      panic("Expected the replaced paragraph");
    }
    const insertedRun = paragraph.content.at(0);
    expect(insertedRun?.type).toBe("run");
    if (insertedRun?.type !== "run") {
      panic("Expected replacement text in a run");
    }
    expect(insertedRun.formatting?.bold).toBe(true);
  });
});

describe("replaceTextInDocument native unit boundaries", () => {
  const nativeUnits = {
    columnBreak: { type: "break", breakType: "column" },
    symbol: { type: "symbol", font: "Symbol", char: "00A7" },
    tab: { type: "tab" },
  } as const satisfies Record<string, Run["content"][number]>;

  const insertAt = (
    unit: Run["content"][number],
    offset: number,
  ): Run["content"] => {
    const source = documentWithContent([
      { type: "run", content: [unit, { type: "text", text: "tail" }] },
    ]);
    const replaced = replaceTextInDocument(
      source,
      {
        start: { paragraphIndex: 0, offset },
        end: { paragraphIndex: 0, offset },
      },
      "X",
    );
    const edited = paragraphOf(replaced).content.at(0);
    if (edited?.type !== "run") panic("Expected the edited run");
    return edited.content;
  };

  test.each(Object.entries(nativeUnits))(
    "inserts after a leading %s",
    (_, unit) => {
      expect(insertAt(unit, 1)).toEqual([
        unit,
        { type: "text", text: "X" },
        { type: "text", text: "tail" },
      ]);
    },
  );

  test.each(Object.entries(nativeUnits))(
    "inserts before a leading %s",
    (_, unit) => {
      expect(insertAt(unit, 0)).toEqual([
        { type: "text", text: "X" },
        unit,
        { type: "text", text: "tail" },
      ]);
    },
  );
});
