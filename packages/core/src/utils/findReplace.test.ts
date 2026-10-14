import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import {
  SOURCE_ZERO_TEXT_BOUNDARY_FIXTURES,
  CONDITIONAL_SOURCE_ZERO_TEXT_BOUNDARY_FIXTURES,
} from "../../../../test/findReplacementBoundaryFixtures";

import type { Document } from "../types/document";
import type { ComplexField, HeaderFooter, Paragraph, Run, SimpleField } from "../types/content";
import { FindReplaceManager } from "../managers/FindReplaceManager";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { findInProseMirrorDocument } from "../prosemirror/findReplaceSelection";
import { resolveDocumentFindRanges } from "./findMatchView";
import { canonicalFieldNode } from "../prosemirror/fieldRepresentation";
import { expectFieldAttrs } from "../prosemirror/attrs";
import {
  createDefaultFindOptions,
  findAllMatches,
  findInDocument,
  getFindParagraphText,
} from "./findReplace";

const document: Document = {
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Outside stock" }] }],
        },
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  content: [
                    {
                      type: "paragraph",
                      content: [{ type: "run", content: [{ type: "text", text: "Inside stock" }] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  },
};

const paragraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const textRun = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

describe("hidden source replacement boundaries", () => {
  const views = [
    { type: "deletion", displayMode: "no-markup" },
    { type: "deletion", displayMode: "simple-markup" },
    { type: "moveFrom", displayMode: "no-markup" },
    { type: "moveFrom", displayMode: "simple-markup" },
    { type: "insertion", displayMode: "original" },
    { type: "moveTo", displayMode: "original" },
  ] as const;
  // Earlier fixtures only hid nonempty text: authored-offset gaps caught those
  // cases, while anchors with no searchable text bypassed that oracle.
  const fixtures = {
    ...SOURCE_ZERO_TEXT_BOUNDARY_FIXTURES,
    ...CONDITIONAL_SOURCE_ZERO_TEXT_BOUNDARY_FIXTURES,
  };
  for (const [kind, content] of Object.entries(fixtures)) {
    test.each(views)(`${kind} in hidden $type ($displayMode) preserves its boundary`, (view) => {
      expect(getFindParagraphText({ type: "paragraph", content })).toBe("");
      const source: Paragraph = {
        type: "paragraph",
        content: [textRun("A"), { type: view.type, info: { id: 1701037 }, content }, textRun("B")],
      };
      const options = { ...createDefaultFindOptions(), displayMode: view.displayMode };
      expect(
        findInDocument({ package: { document: { content: [source] } } }, "AB", options).at(0)
          ?.replacement,
      ).toEqual({ status: "blocked", reason: "hidden-content" });
      for (const text of ["A", "B"]) {
        expect(
          findInDocument({ package: { document: { content: [source] } } }, text, options).at(0)
            ?.replacement,
        ).toEqual({ status: "allowed" });
      }
    });
    test(`${kind} remains a replacement barrier when visible`, () => {
      const source: Paragraph = {
        type: "paragraph",
        content: [textRun("A"), ...content, textRun("B")],
      };
      const matches = findInDocument(
        { package: { document: { content: [source] } } },
        "AB",
        createDefaultFindOptions(),
      );
      if (kind === "pageBreak") expect(matches).toEqual([]);
      else expect(matches.at(0)?.replacement.status).toBe("blocked");
    });
  }
  test.each(views)("empty hidden $type ($displayMode) retains revision ownership", (view) => {
    const source: Paragraph = {
      type: "paragraph",
      content: [
        textRun("A"),
        { type: view.type, info: { id: 1701037 }, content: [] },
        textRun("B"),
      ],
    };
    expect(
      findInDocument({ package: { document: { content: [source] } } }, "AB", {
        ...createDefaultFindOptions(),
        displayMode: view.displayMode,
      }).at(0)?.replacement,
    ).toEqual({ status: "blocked", reason: "hidden-content" });
  });
});

describe("note reference search offsets", () => {
  const cases = (["footnoteRef", "endnoteRef"] as const).flatMap((type) =>
    [1, 12, 123].map((id) => ({ type, id })),
  );

  test.each(cases)(
    "$type id $id does not shift ordinary matches or selected glyphs",
    ({ type, id }) => {
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
                      { type, id },
                      { type: "text", text: "hit hit" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };
      const pmDoc = toProseDoc(source);
      const options = createDefaultFindOptions();
      const matches = findInDocument(source, "hit", options);
      const project = ({
        startOffset,
        endOffset,
        text,
      }: {
        startOffset: number;
        endOffset: number;
        text: string;
      }) => ({ startOffset, endOffset, text });
      expect(matches.map(project)).toEqual([
        { startOffset: 0, endOffset: 3, text: "hit" },
        { startOffset: 4, endOffset: 7, text: "hit" },
      ]);
      expect(findInProseMirrorDocument(pmDoc, "hit", options).map(project)).toEqual(
        matches.map(project),
      );
      const ranges = resolveDocumentFindRanges({ document: source, bodyDoc: pmDoc, matches });
      expect(ranges).toHaveLength(2);
      expect(ranges.map(({ from, to }) => pmDoc.textBetween(from, to))).toEqual(["hit", "hit"]);
      expect(findInProseMirrorDocument(pmDoc, String(id), options)).toHaveLength(0);
    },
  );

  test.each(["footnoteRef", "endnoteRef"] as const)(
    "multiple %s anchors keep all later hits aligned",
    (type) => {
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
                      { type, id: 1 },
                      { type: "text", text: " first hit " },
                      { type, id: 12 },
                      { type: "text", text: " second hit " },
                      { type, id: 123 },
                      { type: "text", text: " third hit" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };
      const pmDoc = toProseDoc(source);
      const options = createDefaultFindOptions();
      const matches = findInDocument(source, "hit", options);
      expect(matches.map(({ startOffset }) => startOffset)).toEqual([7, 19, 30]);
      expect(
        findInProseMirrorDocument(pmDoc, "hit", options).map(({ startOffset }) => startOffset),
      ).toEqual([7, 19, 30]);
      const ranges = resolveDocumentFindRanges({ document: source, bodyDoc: pmDoc, matches });
      expect(ranges).toHaveLength(3);
      expect(ranges.map(({ from, to }) => pmDoc.textBetween(from, to))).toEqual([
        "hit",
        "hit",
        "hit",
      ]);
    },
  );
});

const fieldTokenRun: Run = {
  type: "run",
  content: [
    { type: "text", text: "Cache" },
    { type: "symbol", char: "00A7", font: "Arial" },
    { type: "softHyphen" },
    { type: "noBreakHyphen" },
    { type: "tab" },
    { type: "break", breakType: "textWrapping" },
    { type: "break", breakType: "column" },
    { type: "text", text: "42 HIT" },
  ],
};

const fieldProjectionFixtures = {
  simpleDigit: {
    type: "simpleField",
    fieldType: "DOCVARIABLE",
    instruction: "DOCVARIABLE digit",
    content: [textRun("42")],
  },
  complexDigit: {
    type: "complexField",
    fieldType: "DOCVARIABLE",
    instruction: "DOCVARIABLE digit",
    fieldCode: [],
    fieldResult: [textRun("42")],
  },
  simpleString: {
    type: "simpleField",
    fieldType: "REF",
    instruction: "REF bookmark",
    content: [textRun("Cached HIT")],
  },
  complexString: {
    type: "complexField",
    fieldType: "REF",
    instruction: "REF bookmark",
    fieldCode: [],
    fieldResult: [textRun("Cached HIT")],
  },
  simpleTokens: {
    type: "simpleField",
    fieldType: "REF",
    instruction: "REF bookmark",
    content: [fieldTokenRun],
  },
  complexTokens: {
    type: "complexField",
    fieldType: "REF",
    instruction: "REF bookmark",
    fieldCode: [],
    fieldResult: [fieldTokenRun],
  },
  simpleHyphens: {
    type: "simpleField",
    fieldType: "REF",
    instruction: "REF bookmark",
    content: [
      {
        type: "run",
        content: [
          { type: "text", text: "soft" },
          { type: "softHyphen" },
          { type: "noBreakHyphen" },
          { type: "text", text: "hard" },
        ],
      },
    ],
  },
  complexHyphens: {
    type: "complexField",
    fieldType: "REF",
    instruction: "REF bookmark",
    fieldCode: [],
    fieldResult: [
      {
        type: "run",
        content: [
          { type: "text", text: "soft" },
          { type: "softHyphen" },
          { type: "noBreakHyphen" },
          { type: "text", text: "hard" },
        ],
      },
    ],
  },
  transparent: {
    type: "simpleField",
    fieldType: "REF",
    instruction: "REF bookmark",
    content: [
      {
        type: "inlineWrapper",
        kind: "smartTag",
        element: "tag",
        content: [textRun("Wrapped HIT")],
      },
    ],
  },
  hyperlink: {
    type: "simpleField",
    fieldType: "REF",
    instruction: "REF bookmark",
    content: [
      { type: "hyperlink", href: "https://example.test", children: [textRun("Linked HIT")] },
    ],
  },
  opaqueRun: {
    type: "complexField",
    fieldType: "REF",
    instruction: "REF bookmark",
    fieldCode: [],
    fieldResult: [
      {
        type: "run",
        content: [
          {
            type: "preservedXml",
            xml: "<w:ruby><w:rubyBase><w:r><w:t>Base HIT</w:t></w:r></w:rubyBase></w:ruby>",
            text: "Base HIT",
          },
        ],
      },
    ],
  },
  opaqueInline: {
    type: "simpleField",
    fieldType: "REF",
    instruction: "REF bookmark",
    content: [
      {
        type: "preservedInline",
        xml: "<w:customXml><w:r><w:t>Inline HIT</w:t></w:r></w:customXml>",
        text: "Inline HIT",
      },
    ],
  },
} as const satisfies Record<string, SimpleField | ComplexField>;

describe("page-dependent field search exclusion", () => {
  const pageFieldTypes = ["PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF"] as const;
  const fixtures = pageFieldTypes.flatMap((fieldType) =>
    ["simpleField", "complexField"].flatMap((type) =>
      (
        [
          {
            name: "plain",
            instruction: `${fieldType}${fieldType === "PAGEREF" ? " target" : ""}`,
            parsedType: fieldType,
          },
          {
            name: "switches",
            instruction: `${fieldType}${fieldType === "PAGEREF" ? " target \\h" : ""} \\* MERGEFORMAT`,
            parsedType: fieldType,
          },
          {
            name: "lowercase",
            instruction: `  ${fieldType.toLowerCase()}${fieldType === "PAGEREF" ? " target" : ""} \\* mergeformat  `,
            parsedType: fieldType,
          },
          {
            name: "unclassified",
            instruction: `  ${fieldType.toLowerCase()}${fieldType === "PAGEREF" ? " target \\h" : ""} \\* MERGEFORMAT `,
            parsedType: "UNKNOWN",
          },
        ] as const satisfies readonly {
          name: string;
          instruction: string;
          parsedType: SimpleField["fieldType"];
        }[]
      ).map(({ name, instruction, parsedType }) => ({
        name: `${type}/${fieldType}/${name}`,
        type,
        instruction,
        parsedType,
      })),
    ),
  );

  test.each(fixtures)(
    "$name excludes cached page digits without shifting later matches",
    ({ type, instruction, parsedType }) => {
      const makeField = (
        fieldType: SimpleField["fieldType"],
        fieldInstruction: string,
      ): SimpleField | ComplexField => {
        if (type === "simpleField")
          return {
            type: "simpleField",
            fieldType,
            instruction: fieldInstruction,
            content: [textRun("3")],
          };
        return {
          type: "complexField",
          fieldType,
          instruction: fieldInstruction,
          fieldCode: [],
          fieldResult: [textRun("3")],
        };
      };
      const source: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  textRun("head "),
                  makeField(parsedType, instruction),
                  textRun(" "),
                  makeField("DOCVARIABLE", "DOCVARIABLE value \\* MERGEFORMAT"),
                  textRun(" tail 3"),
                ],
              },
            ],
          },
        },
      };
      const pmDoc = toProseDoc(source);
      const pageOnly: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [makeField(parsedType, instruction)],
              },
            ],
          },
        },
      };
      const pageOnlyPm = toProseDoc(pageOnly);
      for (const displayMode of [undefined, "all-markup", "no-markup"] as const) {
        const options = {
          ...createDefaultFindOptions(),
          ...(displayMode === undefined ? {} : { displayMode }),
        };
        const project = ({
          paragraphIndex,
          startOffset,
          endOffset,
          text,
        }: {
          paragraphIndex: number;
          startOffset: number;
          endOffset: number;
          text: string;
        }) => ({ paragraphIndex, startOffset, endOffset, text });
        const expected = [
          { paragraphIndex: 0, startOffset: 7, endOffset: 8, text: "3" },
          { paragraphIndex: 0, startOffset: 14, endOffset: 15, text: "3" },
        ];
        expect(findInDocument(source, "head 3", options)).toEqual([]);
        expect(findInProseMirrorDocument(pmDoc, "head 3", options)).toEqual([]);
        const onlyPageField: Document = {
          package: {
            document: {
              content: [{ type: "paragraph", content: [makeField(parsedType, instruction)] }],
            },
          },
        };
        expect(findInDocument(onlyPageField, "3", options)).toEqual([]);
        expect(findInDocument(source, "3", options).map(project)).toEqual(expected);
        expect(findInProseMirrorDocument(pmDoc, "3", options).map(project)).toEqual(expected);
        expect(findInDocument(source, "head 3", options)).toHaveLength(0);
        expect(findInProseMirrorDocument(pmDoc, "head 3", options)).toHaveLength(0);
        expect(findInDocument(pageOnly, "3", options)).toHaveLength(0);
        expect(findInProseMirrorDocument(pageOnlyPm, "3", options)).toHaveLength(0);
        expect(findInDocument(source, "tail", options).map(project)).toEqual([
          { paragraphIndex: 0, startOffset: 9, endOffset: 13, text: "tail" },
        ]);
      }
    },
  );
});

describe("field and native text search projection parity", () => {
  test.each(Object.entries(fieldProjectionFixtures))(
    "%s preserves cached tokens, substring hits, and trailing offsets through conversion",
    (name, field) => {
      const sourceParagraph: Paragraph = {
        type: "paragraph",
        content: [textRun("before "), field, textRun(" tail HIT hit")],
      };
      const source: Document = { package: { document: { content: [sourceParagraph] } } };
      const pmDoc = toProseDoc(source);
      const fieldNode = pmDoc.child(0).child(1);
      if (fieldNode.type.name === "structuredField") {
        const normalized = canonicalFieldNode(fieldNode);
        expect(expectFieldAttrs(normalized).displayText).toBe(
          getFindParagraphText({ type: "paragraph", content: [field] }),
        );
        expect(canonicalFieldNode(normalized)).toBe(normalized);
      }
      const saved = fromProseDoc(pmDoc, source, { stylesheetSource: { type: "package" } });
      const savedParagraph = saved.package.document.content.at(0);
      if (savedParagraph?.type !== "paragraph") panic("Roundtrip paragraph missing");
      const visibleText = getFindParagraphText(sourceParagraph);
      expect(getFindParagraphText(savedParagraph)).toBe(visibleText);
      const queries = new Set([...visibleText, "tail", "HIT", "hit", "42", "Cached", "before "]);
      for (const options of [
        createDefaultFindOptions(),
        { matchCase: true, matchWholeWord: false },
      ]) {
        for (const query of queries) {
          const project = ({
            paragraphIndex,
            startOffset,
            endOffset,
            text,
          }: {
            paragraphIndex: number;
            startOffset: number;
            endOffset: number;
            text: string;
          }) => ({ paragraphIndex, startOffset, endOffset, text });
          const modelMatches = findInDocument(source, query, options).map(project);
          expect(findInProseMirrorDocument(pmDoc, query, options).map(project)).toEqual(
            modelMatches,
          );
          expect(findInDocument(saved, query, options).map(project)).toEqual(modelMatches);
          if (query === "tail") expect(modelMatches).toHaveLength(1);
        }
      }
      expect(findInDocument(source, "REF bookmark", createDefaultFindOptions())).toHaveLength(0);
      if (
        [
          "simpleDigit",
          "complexDigit",
          "simpleString",
          "complexString",
          "simpleHyphens",
          "complexHyphens",
          "complexTokens",
          "opaqueRun",
        ].includes(name)
      ) {
        expect(pmDoc.child(0).child(1).type.name).toBe("field");
      }
      if (["simpleTokens", "transparent", "hyperlink", "opaqueInline"].includes(name)) {
        expect(pmDoc.child(0).child(1).type.name).toBe("structuredField");
      }
    },
  );
});

const allStoriesDocument = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("body hit"),
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [
                {
                  type: "shape",
                  shape: {
                    type: "shape",
                    shapeType: "textBox",
                    size: { width: 914400, height: 914400 },
                    textBody: { content: [paragraph("box hit")] },
                  },
                },
              ],
            },
          ],
        },
        paragraph("last body hit"),
      ],
      sections: [
        {
          properties: {
            headerReferences: [
              { type: "even", rId: "he" },
              { type: "first", rId: "hf" },
              { type: "default", rId: "hd" },
            ],
            footerReferences: [
              { type: "even", rId: "fe" },
              { type: "first", rId: "ff" },
              { type: "default", rId: "fd" },
            ],
          },
          content: [],
        },
      ],
    },
    headers: new Map(
      ["he", "hf", "hd"].map((rId) => [
        rId,
        {
          type: "header",
          hdrFtrType: "default",
          content: [paragraph(`${rId} hit`)],
        } as const satisfies HeaderFooter,
      ]),
    ),
    footers: new Map(
      ["fe", "ff", "fd"].map((rId) => [
        rId,
        {
          type: "footer",
          hdrFtrType: "default",
          content: [paragraph(`${rId} hit`)],
        } as const satisfies HeaderFooter,
      ]),
    ),
    footnotes: [
      { type: "footnote", id: -1, noteType: "separator", content: [paragraph("separator hit")] },
      { type: "footnote", id: 1, content: [paragraph("footnote hit")] },
    ],
    endnotes: [{ type: "endnote", id: 2, content: [paragraph("endnote hit")] }],
  },
});

describe("document-wide story search", () => {
  test("covers each story and variant once, with stable navigation order and shape addresses", () => {
    const doc = allStoriesDocument();
    const before = structuredClone(doc);
    const matches = findInDocument(doc, "hit", createDefaultFindOptions());
    expect(matches).toHaveLength(11);
    expect(matches.map(({ story, textBoxIndex }) => ({ story, textBoxIndex }))).toEqual([
      { story: "main", textBoxIndex: undefined },
      { story: "main", textBoxIndex: 0 },
      { story: "main", textBoxIndex: undefined },
      ...["hd", "hf", "he"].map((rId) => ({
        story: { kind: "header", rId },
        textBoxIndex: undefined,
      })),
      ...["fd", "ff", "fe"].map((rId) => ({
        story: { kind: "footer", rId },
        textBoxIndex: undefined,
      })),
      { story: { kind: "footnote", id: 1 }, textBoxIndex: undefined },
      { story: { kind: "endnote", id: 2 }, textBoxIndex: undefined },
    ]);
    expect(matches.at(1)).toMatchObject({
      paragraphIndex: 0,
      anchorParagraphIndex: 1,
      textBoxIndex: 0,
    });
    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    for (const match of matches.slice(1)) expect(manager.navigate("next")?.match).toEqual(match);
    expect(manager.navigate("next")?.match).toEqual(matches.at(0));
    expect(manager.navigate("previous")?.match).toEqual(matches.at(-1));
    expect(doc).toEqual(before);
  });

  test.each(["all-markup", "simple-markup", "no-markup", "original"] as const)(
    "%s respects revision visibility without shifting authored offsets",
    (displayMode) => {
      const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });
      const doc: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  run("before "),
                  {
                    type: "deletion",
                    info: { id: 1, author: "Author" },
                    content: [run("deleted hit ")],
                  },
                  {
                    type: "insertion",
                    info: { id: 2, author: "Author" },
                    content: [run("inserted hit ")],
                  },
                  run("after hit"),
                ],
              },
            ],
          },
        },
      };
      const options = { ...createDefaultFindOptions(), displayMode };
      const matches = findInDocument(doc, "hit", options);
      expect(matches).toHaveLength(displayMode === "all-markup" ? 3 : 2);
      expect(matches.at(-1)?.startOffset).toBe(38);
      expect(findInDocument(doc, "deleted", options)).toHaveLength(
        displayMode === "original" || displayMode === "all-markup" ? 1 : 0,
      );
      expect(findInDocument(doc, "inserted", options)).toHaveLength(
        displayMode === "original" ? 0 : 1,
      );
    },
  );

  test("secondary-story replacement cannot target a coincident body paragraph", () => {
    const doc = allStoriesDocument();
    const matches = findInDocument(doc, "hit", createDefaultFindOptions());
    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    for (const [index, match] of matches.entries()) {
      if (match.story === "main" && match.textBoxIndex === undefined) continue;
      expect(manager.goTo(index)?.match).toEqual(match);
      expect(manager.replaceCurrent(doc, "changed")).toBeNull();
    }
    expect(manager.replaceAll(doc, matches, "changed")).toMatchObject({
      replacedCount: 2,
      skippedCount: 9,
    });
    expect(findInDocument(doc, "hit", createDefaultFindOptions())).toHaveLength(11);
  });

  test("only rendered variants count when the view supplies story eligibility", () => {
    const matches = findInDocument(allStoriesDocument(), "hit", createDefaultFindOptions(), {
      renderedStories: [
        "main",
        { kind: "header", rId: "hd" },
        { kind: "header", rId: "hd" },
        { kind: "footer", rId: "fd" },
      ],
    });
    expect(matches).toHaveLength(7);
    expect(matches.map(({ story }) => story)).toEqual([
      "main",
      "main",
      "main",
      { kind: "header", rId: "hd" },
      { kind: "footer", rId: "fd" },
      { kind: "footnote", id: 1 },
      { kind: "endnote", id: 2 },
    ]);
  });

  test("every story keeps case and whole-word behavior", () => {
    const doc = allStoriesDocument();
    expect(findInDocument(doc, "HIT", { matchCase: false, matchWholeWord: true })).toHaveLength(11);
    expect(findInDocument(doc, "HIT", { matchCase: true, matchWholeWord: true })).toHaveLength(0);
    expect(findInDocument(doc, "hi", { matchCase: false, matchWholeWord: true })).toHaveLength(0);
    expect(findInDocument(doc, "hi", { matchCase: false, matchWholeWord: false })).toHaveLength(11);
  });

  test("hidden shapes retain authored numbering for later visible shape addresses", () => {
    const doc = allStoriesDocument();
    const shapeAnchor = doc.package.document.content.at(1);
    if (shapeAnchor?.type !== "paragraph") panic("Fixture shape anchor missing");
    const shapeRun = shapeAnchor.content.at(0);
    if (shapeRun?.type !== "run") panic("Fixture shape run missing");
    doc.package.document.content = [
      {
        type: "paragraph",
        content: [
          { type: "deletion", info: { id: 1, author: "Author" }, content: [shapeRun] },
          shapeRun,
        ],
      },
    ];
    const matches = findInDocument(doc, "box", {
      ...createDefaultFindOptions(),
      displayMode: "no-markup",
    });
    expect(matches).toMatchObject([
      { story: "main", textBoxIndex: 1, anchorParagraphIndex: 0, paragraphIndex: 0 },
    ]);
  });

  test("comment text is opt-in, has its own address, and cannot replace body text", () => {
    const doc = allStoriesDocument();
    doc.package.document.comments = [
      { id: 7, author: "Author", content: [paragraph("comment hit")] },
    ];
    expect(findInDocument(doc, "comment", createDefaultFindOptions())).toHaveLength(0);
    const matches = findInDocument(doc, "comment", {
      ...createDefaultFindOptions(),
      includeComments: true,
    });
    expect(matches).toMatchObject([
      { story: "main", commentId: 7, paragraphIndex: 0, startOffset: 0 },
    ]);
    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    expect(manager.replaceCurrent(doc, "changed")).toBeNull();
    expect(manager.replaceAll(doc, matches, "changed")).toEqual({
      document: doc,
      replacedCount: 0,
      skippedCount: 1,
    });
  });

  test.each(["all-markup", "simple-markup", "no-markup", "original"] as const)(
    "%s respects table row and cell visibility and keeps authored paragraph indices",
    (displayMode) => {
      const doc: Document = {
        package: {
          document: {
            content: [
              {
                type: "table",
                rows: [
                  {
                    type: "tableRow",
                    structuralChange: {
                      type: "tableRowDeletion",
                      info: { id: 1, author: "Author" },
                    },
                    cells: [{ type: "tableCell", content: [paragraph("deleted row hit")] }],
                  },
                  {
                    type: "tableRow",
                    cells: [
                      {
                        type: "tableCell",
                        structuralChange: {
                          type: "tableCellInsertion",
                          info: { id: 2, author: "Author" },
                        },
                        content: [paragraph("inserted cell hit")],
                      },
                      { type: "tableCell", content: [paragraph("visible hit")] },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };
      const matches = findInDocument(doc, "hit", { ...createDefaultFindOptions(), displayMode });
      const expected = {
        "all-markup": [0, 1, 2],
        original: [0, 2],
        "no-markup": [1, 2],
        "simple-markup": [1, 2],
      };
      expect(matches.map(({ paragraphIndex }) => paragraphIndex)).toEqual(expected[displayMode]);
    },
  );
});

describe("regex matching", () => {
  test("honors regex mode and groups whole-word alternation", () => {
    expect(
      findAllMatches("hit 42 hit 99", "hit \\d+", {
        matchCase: true,
        matchWholeWord: false,
        useRegex: true,
      }),
    ).toEqual([
      { start: 0, end: 6 },
      { start: 7, end: 13 },
    ]);
    expect(
      findAllMatches("foox xbar foo bar", "foo|bar", {
        matchCase: false,
        matchWholeWord: true,
        useRegex: true,
      }),
    ).toEqual([
      { start: 10, end: 13 },
      { start: 14, end: 17 },
    ]);
    expect(
      findAllMatches("hit", "[", { matchCase: false, matchWholeWord: false, useRegex: true }),
    ).toEqual([]);
  });

  test("skips empty regex hits and advances over complete Unicode code points", () => {
    const options = { matchCase: false, matchWholeWord: false, useRegex: true };
    expect(findAllMatches("📝 hit 📝 hit", "(?=📝)|hit", options)).toEqual([
      { start: 3, end: 6 },
      { start: 10, end: 13 },
    ]);
    for (const pattern of ["^", "$", "a*", "(?=.)", "(?:)"]) {
      expect(findAllMatches("📝 xyz", pattern, options)).toEqual([]);
    }
  });
});

describe("findInDocument", () => {
  test("finds matches in body and table paragraphs with stable paragraph offsets", () => {
    const matches = findInDocument(document, "stock", createDefaultFindOptions());

    expect(
      matches.map(({ paragraphIndex, startOffset, text }) => ({
        paragraphIndex,
        startOffset,
        text,
      })),
    ).toEqual([
      { paragraphIndex: 0, startOffset: 8, text: "stock" },
      { paragraphIndex: 1, startOffset: 7, text: "stock" },
    ]);
  });

  test("honors case and whole-word options", () => {
    const caseDocument: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Stock stock stockholder" }],
                },
              ],
            },
          ],
        },
      },
    };

    expect(
      findInDocument(caseDocument, "Stock", { matchCase: true, matchWholeWord: true }),
    ).toHaveLength(1);
    expect(
      findInDocument(caseDocument, "stock", { matchCase: false, matchWholeWord: true }),
    ).toHaveLength(2);
  });

  test("symbol run children contribute visible characters and preserve later offsets", () => {
    const doc: Document = {
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
                    { type: "symbol", char: "00A7", font: "Arial" },
                    { type: "text", text: "B" },
                  ],
                },
              ],
            },
          ],
        },
      },
    };
    expect(findInDocument(doc, "A§B", createDefaultFindOptions())).toMatchObject([
      { startOffset: 0, endOffset: 3 },
    ]);
    expect(findInDocument(doc, "B", createDefaultFindOptions())).toMatchObject([
      { startOffset: 2, endOffset: 3 },
    ]);
    const pmDoc = toProseDoc(doc);
    for (const query of ["A§B", "§", "B"]) {
      const modelMatches = findInDocument(doc, query, createDefaultFindOptions());
      const pmMatches = findInProseMirrorDocument(pmDoc, query, createDefaultFindOptions());
      expect(
        modelMatches.map(({ paragraphIndex, startOffset, endOffset, text }) => ({
          paragraphIndex,
          startOffset,
          endOffset,
          text,
        })),
      ).toEqual(
        pmMatches.map(({ paragraphIndex, startOffset, endOffset, text }) => ({
          paragraphIndex,
          startOffset,
          endOffset,
          text,
        })),
      );
    }
  });

  test("does not match text across an explicit page-break run", () => {
    const pageBreakDocument: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    { type: "text", text: "left" },
                    { type: "break", breakType: "page" },
                    { type: "text", text: "right" },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    expect(findInDocument(pageBreakDocument, "left", createDefaultFindOptions())).toHaveLength(1);
    expect(findInDocument(pageBreakDocument, "right", createDefaultFindOptions())).toHaveLength(1);
    expect(findInDocument(pageBreakDocument, "leftright", createDefaultFindOptions())).toEqual([]);
  });

  test.each([undefined, "textWrapping"] as const)(
    "counts a %s line-break type in search offsets",
    (breakType) => {
      const lineBreakDocument: Document = {
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
                      {
                        type: "break",
                        ...(breakType !== undefined ? { breakType } : {}),
                      },
                      { type: "text", text: "B" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };

      expect(findInDocument(lineBreakDocument, "A\nB", createDefaultFindOptions())).toHaveLength(1);
      expect(
        findInDocument(lineBreakDocument, "B", createDefaultFindOptions()).at(0),
      ).toMatchObject({ startOffset: 2 });
    },
  );
});
