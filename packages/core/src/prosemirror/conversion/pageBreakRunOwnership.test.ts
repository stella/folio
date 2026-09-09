import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { Transform } from "prosemirror-transform";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import type {
  Document,
  Hyperlink,
  Paragraph,
  RunContent,
  RunPropertyChange,
  TrackedRunChange,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const OUTER_REVISION = {
  id: 91,
  author: "Reviewer",
  date: "2026-09-09T00:00:00.000Z",
} as const;

const PROPERTY_REVISION = {
  type: "runPropertyChange",
  info: {
    id: 92,
    author: "Reviewer",
    date: "2026-09-09T00:01:00.000Z",
  },
  previousFormatting: { italic: true },
  currentFormatting: { underline: { style: "single" } },
} as const satisfies RunPropertyChange;

const SHAPES = {
  leading: [
    { type: "break", breakType: "page", clear: "all" },
    { type: "text", text: "A" },
  ],
  trailing: [
    { type: "text", text: "A" },
    { type: "break", breakType: "page", clear: "all" },
  ],
  interior: [
    { type: "text", text: "A" },
    { type: "break", breakType: "page", clear: "all" },
    { type: "text", text: "B" },
  ],
  multiple: [
    { type: "break", breakType: "page", clear: "all" },
    { type: "text", text: "A" },
    { type: "break", breakType: "page", clear: "all" },
    { type: "text", text: "B" },
    { type: "break", breakType: "page", clear: "all" },
  ],
} as const satisfies Record<string, readonly RunContent[]>;

const WRAPPERS = ["insertion", "deletion", "moveFrom", "moveTo"] as const;

const NON_TEXT_CARRIERS = {
  drawing: {
    type: "drawing",
    image: {
      type: "image",
      rId: "rIdImage1",
      src: "data:image/png;base64,AA==",
      size: { width: 914_400, height: 457_200 },
      wrap: { type: "inline" },
    },
  },
  shape: {
    type: "shape",
    shape: {
      type: "shape",
      shapeType: "rect",
      size: { width: 914_400, height: 457_200 },
      wrap: { type: "inline" },
    },
  },
  footnoteRef: { type: "footnoteRef", id: 3 },
  endnoteRef: { type: "endnoteRef", id: 4 },
} as const satisfies Record<string, RunContent>;

const UNREPRESENTABLE_PAGE_BREAK_SIBLINGS = {
  fieldChar: { type: "fieldChar", charType: "begin" },
  instrText: { type: "instrText", text: " PAGE " },
  noBreakHyphen: { type: "noBreakHyphen" },
  softHyphen: { type: "softHyphen" },
  "text-box shape": {
    type: "shape",
    shape: {
      type: "shape",
      shapeType: "rect",
      size: { width: 914_400, height: 457_200 },
      textBody: { content: [] },
    },
  },
} as const satisfies Record<string, RunContent>;

const unrepresentablePageBreakSiblingMessage = (sibling: RunContent): string =>
  sibling.type === "shape"
    ? "A page-break-bearing run containing a text-box shape cannot be represented in the editor model"
    : `A page-break-bearing run containing ${sibling.type} cannot be represented in the editor model`;

const trackedDocument = (
  type: (typeof WRAPPERS)[number],
  contents: readonly RunContent[],
  withPropertyChange = true,
): Document => {
  const document = createEmptyDocument();
  const change: TrackedRunChange = {
    type,
    info: OUTER_REVISION,
    content: [
      {
        type: "run",
        formatting: { underline: { style: "single" } },
        ...(withPropertyChange ? { propertyChanges: [PROPERTY_REVISION] } : {}),
        content: [...contents],
      },
    ],
  };
  document.package.document.content = [{ type: "paragraph", content: [change] }];
  return document;
};

const firstTrackedRun = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected paragraph");
  }
  const change = paragraph.content.at(0);
  if (
    change?.type !== "insertion" &&
    change?.type !== "deletion" &&
    change?.type !== "moveFrom" &&
    change?.type !== "moveTo"
  ) {
    throw new Error("Expected tracked run wrapper");
  }
  const run = change.content.at(0);
  if (run?.type !== "run") {
    throw new Error("Expected run");
  }
  return { change, run };
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/document.xml");
  if (!file) {
    throw new Error("Missing word/document.xml");
  }
  return file.async("text");
};

const pageBreakCount = (xml: string): number =>
  xml.match(/<w:br\b[^>]*\bw:type="page"/gu)?.length ?? 0;

const cleanParagraphProjection = (paragraph: Paragraph): string[] => {
  const projection: string[] = [];
  for (const item of paragraph.content) {
    if (item.type !== "run") {
      continue;
    }
    for (const content of item.content) {
      if (content.type === "text") {
        projection.push(content.text);
      } else if (content.type === "break" && content.breakType === "page") {
        projection.push("<page>");
      }
    }
  }
  return projection;
};

describe("page-break run ownership", () => {
  test.each([undefined, "none", "left", "right", "all"] as const)(
    "writes an explicit page type and preserves clear=%s",
    async (clear) => {
      const source = createEmptyDocument();
      source.package.document.content = [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [
                {
                  type: "break",
                  breakType: "page",
                  ...(clear !== undefined ? { clear } : {}),
                },
              ],
            },
          ],
        },
      ];

      const prose = toProseDoc(source);
      expect(prose.firstChild?.firstChild?.type.name).toBe("pageBreakRun");
      expect(prose.firstChild?.firstChild?.attrs["clear"] ?? undefined).toBe(clear);
      const roundTripped = fromProseDoc(prose, source);
      const paragraph = roundTripped.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") {
        throw new Error("Expected paragraph");
      }
      const run = paragraph.content.at(0);
      if (run?.type !== "run") {
        throw new Error("Expected run");
      }
      expect(run.content).toEqual([
        {
          type: "break",
          breakType: "page",
          ...(clear !== undefined ? { clear } : {}),
        },
      ]);

      const xml = await documentXml(await createDocx(roundTripped));
      expect(xml).toContain('w:type="page"');
      expect(xml.includes("w:clear=")).toBe(clear !== undefined);
      if (clear !== undefined) {
        expect(xml).toContain(`w:clear="${clear}"`);
      }
    },
  );

  test("rejoins every serializable sibling from one page-break-bearing run", () => {
    const source = createEmptyDocument();
    const mixedContent = [
      { type: "text", text: "A" },
      { type: "tab" },
      { type: "break" },
      { type: "break", breakType: "textWrapping", clear: "left" },
      { type: "break", breakType: "column", clear: "right" },
      { type: "break", breakType: "page", clear: "right" },
      { type: "renderedPageBreak" },
      { type: "symbol", font: "Wingdings", char: "F06F" },
      { type: "text", text: "B" },
    ] as const satisfies readonly RunContent[];
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            formatting: { underline: { style: "single" } },
            propertyChanges: [PROPERTY_REVISION],
            content: [...mixedContent],
          },
        ],
      },
    ];

    const prose = toProseDoc(source);
    const cloned = prose.type.schema.nodeFromJSON(prose.toJSON());
    const restored = fromProseDoc(cloned, source);
    const paragraph = restored.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }

    expect(paragraph.content).toHaveLength(1);
    const run = paragraph.content.at(0);
    if (run?.type !== "run") {
      throw new Error("Expected run");
    }
    expect(run.content).toEqual(mixedContent);
    expect(run.formatting?.underline).toEqual({ style: "single" });
    expect(run.propertyChanges).toEqual([PROPERTY_REVISION]);
  });

  test.each(Object.entries(NON_TEXT_CARRIERS))(
    "rejoins a %s with its same-source-run page break",
    (_, carrier) => {
      const source = createEmptyDocument();
      source.package.document.content = [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              formatting: { bold: true },
              propertyChanges: [PROPERTY_REVISION],
              content: [carrier, { type: "break", breakType: "page" }, { type: "text", text: "A" }],
            },
          ],
        },
      ];

      const prose = toProseDoc(source);
      const restored = fromProseDoc(prose.type.schema.nodeFromJSON(prose.toJSON()), source);
      const paragraph = restored.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") {
        throw new Error("Expected paragraph");
      }
      expect(paragraph.content).toHaveLength(1);
      const run = paragraph.content.at(0);
      if (run?.type !== "run") {
        throw new Error("Expected run");
      }
      expect(run.content.map(({ type }) => type)).toEqual([carrier.type, "break", "text"]);
      expect(run.formatting).toEqual({ bold: true });
      expect(run.propertyChanges).toEqual([PROPERTY_REVISION]);
    },
  );

  test("preserves sparse direct formatting without materializing inherited formatting", () => {
    const source = createEmptyDocument();
    source.package.styles = {
      docDefaults: {
        rPr: {
          bold: true,
          boldCs: true,
          italic: true,
          italicCs: true,
          fontSizeCs: 26,
        },
      },
      styles: [],
    };
    const directFormatting = {
      bold: false,
      boldCs: false,
      italicCs: false,
      fontSize: 22,
      underline: { style: "none" },
    } as const;
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            formatting: directFormatting,
            content: [{ type: "break", breakType: "page" }],
          },
        ],
      },
    ];

    const prose = toProseDoc(source, { styles: source.package.styles });
    const restored = fromProseDoc(prose.type.schema.nodeFromJSON(prose.toJSON()), source);
    const paragraph = restored.package.document.content.at(0);
    const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
    expect(run?.type).toBe("run");
    if (run?.type === "run") {
      expect(run.formatting).toEqual(directFormatting);
    }
  });

  test("uses the source-run owner id as the join invariant", () => {
    const source = createEmptyDocument();
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            content: [
              { type: "text", text: "A" },
              { type: "break", breakType: "page" },
              { type: "text", text: "B" },
            ],
          },
        ],
      },
    ];

    const prose = toProseDoc(source);
    const ownerType = prose.type.schema.marks["pageBreakRunOwner"];
    if (!ownerType) {
      throw new Error("Expected page-break source-run owner mark");
    }
    let trailingTextPosition: number | undefined;
    prose.descendants((node, position) => {
      if (node.isText && node.text === "B") {
        trailingTextPosition = position;
      }
    });
    if (trailingTextPosition === undefined) {
      throw new Error("Expected trailing text");
    }
    const mutated = new Transform(prose)
      .removeMark(trailingTextPosition, trailingTextPosition + 1, ownerType)
      .addMark(trailingTextPosition, trailingTextPosition + 1, ownerType.create({ id: 99 })).doc;
    const paragraph = fromProseDoc(mutated, source).package.document.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }

    expect(paragraph.content).toHaveLength(2);
    expect(paragraph.content.map((item) => (item.type === "run" ? item.content : []))).toEqual([
      [
        { type: "text", text: "A" },
        { type: "break", breakType: "page" },
      ],
      [{ type: "text", text: "B" }],
    ]);
  });

  test("keeps non-page break type provenance and clear values at a package fixed point", async () => {
    const source = createEmptyDocument();
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            content: [
              { type: "break" },
              { type: "break", breakType: "textWrapping", clear: "left" },
              { type: "break", breakType: "column", clear: "right" },
              { type: "break", breakType: "page", clear: "all" },
            ],
          },
        ],
      },
    ];

    const first = await createDocx(fromProseDoc(toProseDoc(source), source));
    const firstXml = await documentXml(first);
    expect(firstXml).toContain("<w:br/>");
    expect(firstXml).toContain('<w:br w:type="textWrapping" w:clear="left"/>');
    expect(firstXml).toContain('<w:br w:type="column" w:clear="right"/>');
    expect(firstXml).toContain('<w:br w:type="page" w:clear="all"/>');

    const reopened = await parseDocx(first);
    const second = await createDocx(fromProseDoc(toProseDoc(reopened), reopened));
    expect(await documentXml(second)).toBe(firstXml);
  });

  test.each(Object.entries(UNREPRESENTABLE_PAGE_BREAK_SIBLINGS))(
    "fails closed for an unrepresentable %s sibling",
    (_, sibling) => {
      const source = createEmptyDocument();
      source.package.document.content = [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [sibling, { type: "break", breakType: "page" }],
            },
          ],
        },
      ];

      expect(() => toProseDoc(source)).toThrow(unrepresentablePageBreakSiblingMessage(sibling));
    },
  );

  for (const wrapper of ["hyperlink", "tracked hyperlink"] as const) {
    test.each(Object.entries(UNREPRESENTABLE_PAGE_BREAK_SIBLINGS))(
      `fails closed for an unrepresentable %s sibling in a ${wrapper}`,
      (_, sibling) => {
        const source = createEmptyDocument();
        const hyperlink: Hyperlink = {
          type: "hyperlink",
          href: "https://example.test/page-break",
          children: [
            {
              type: "run",
              content: [sibling, { type: "break", breakType: "page" }],
            },
          ],
        };
        source.package.document.content = [
          {
            type: "paragraph",
            content: [
              wrapper === "hyperlink"
                ? hyperlink
                : { type: "insertion", info: OUTER_REVISION, content: [hyperlink] },
            ],
          },
        ];

        expect(() => toProseDoc(source)).toThrow(unrepresentablePageBreakSiblingMessage(sibling));
      },
    );
  }

  test("does not merge equal-format runs across a page-break source-run boundary", () => {
    const source = createEmptyDocument();
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "A" }] },
          {
            type: "run",
            formatting: { bold: true },
            content: [{ type: "break", breakType: "page" }],
          },
          { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "B" }] },
        ],
      },
    ];

    const prose = toProseDoc(source);
    const restored = fromProseDoc(prose.type.schema.nodeFromJSON(prose.toJSON()), source);
    const paragraph = restored.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }

    expect(paragraph.content).toHaveLength(3);
    expect(paragraph.content.map((item) => (item.type === "run" ? item.content : []))).toEqual([
      [{ type: "text", text: "A" }],
      [{ type: "break", breakType: "page" }],
      [{ type: "text", text: "B" }],
    ]);
  });

  test("keeps ordinary equal-format text run coalescing unchanged", () => {
    const source = createEmptyDocument();
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "A" }] },
          { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "B" }] },
        ],
      },
    ];

    const restored = fromProseDoc(toProseDoc(source), source);
    const paragraph = restored.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }

    expect(paragraph.content).toHaveLength(1);
    expect(paragraph.content.at(0)?.type).toBe("run");
    if (paragraph.content.at(0)?.type === "run") {
      expect(paragraph.content.at(0)?.content).toEqual([{ type: "text", text: "AB" }]);
    }
  });

  for (const type of WRAPPERS) {
    for (const [shape, contents] of Object.entries(SHAPES)) {
      test(`preserves ${shape} and same-run adjacency in a ${type}`, () => {
        const source = trackedDocument(type, contents);
        const prose = toProseDoc(source);
        const cloned = prose.type.schema.nodeFromJSON(prose.toJSON());
        const pageBreaks: (typeof prose)[] = [];
        cloned.descendants((node) => {
          if (node.type.name === "pageBreakRun") {
            pageBreaks.push(node);
          }
        });

        expect(pageBreaks).toHaveLength(
          contents.filter((content) => content.type === "break" && content.breakType === "page")
            .length,
        );
        expect(pageBreaks.every((node) => node.attrs["clear"] === "all")).toBe(true);
        expect(
          pageBreaks.every((node) =>
            node.marks.some(
              (mark) =>
                mark.type.name ===
                (type === "deletion" || type === "moveFrom" ? "deletion" : "insertion"),
            ),
          ),
        ).toBe(true);
        expect(
          pageBreaks.every((node) =>
            node.marks.some((mark) => mark.type.name === "runPropertyChange"),
          ),
        ).toBe(true);

        const { change, run } = firstTrackedRun(fromProseDoc(cloned, source));
        expect(change.type).toBe(type);
        expect(run.content).toEqual(contents);
        expect(run.formatting?.underline).toEqual({ style: "single" });
        expect(run.propertyChanges).toEqual([PROPERTY_REVISION]);
      });
    }
  }

  test("preserves ownership through nested content-control, revision, and hyperlink wrappers", () => {
    const source = createEmptyDocument();
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type: "inlineSdt",
            properties: { sdtType: "richText", alias: "owned break" },
            content: [
              {
                type: "insertion",
                info: OUTER_REVISION,
                content: [
                  {
                    type: "hyperlink",
                    href: "https://example.test/owned-break",
                    children: [
                      {
                        type: "run",
                        formatting: { underline: { style: "single" } },
                        propertyChanges: [PROPERTY_REVISION],
                        content: [...SHAPES.interior],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    ];

    const prose = toProseDoc(source);
    const pageBreaks: (typeof prose)[] = [];
    prose.descendants((node) => {
      if (node.type.name === "pageBreakRun") {
        pageBreaks.push(node);
      }
    });
    expect(pageBreaks).toHaveLength(1);
    expect(pageBreaks.at(0)?.marks.map(({ type }) => type.name)).toEqual(
      expect.arrayContaining(["hyperlink", "insertion", "runPropertyChange"]),
    );

    const restored = fromProseDoc(prose, source);
    const paragraph = restored.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }
    const sdt = paragraph.content.at(0);
    if (sdt?.type !== "inlineSdt") {
      throw new Error("Expected inline content control");
    }
    const insertion = sdt.content.at(0);
    if (insertion?.type !== "insertion") {
      throw new Error("Expected insertion");
    }
    const hyperlink = insertion.content.at(0);
    if (hyperlink?.type !== "hyperlink") {
      throw new Error("Expected hyperlink");
    }
    const run = hyperlink.children.at(0);
    if (run?.type !== "run") {
      throw new Error("Expected run");
    }
    expect(hyperlink.href).toBe("https://example.test/owned-break");
    expect(run.content).toEqual(SHAPES.interior);
    expect(run.formatting?.underline).toEqual({ style: "single" });
    expect(run.propertyChanges).toEqual([PROPERTY_REVISION]);
  });

  test.each(WRAPPERS)("enumerates and resolves a stacked %s plus rPrChange", async (type) => {
    const source = trackedDocument(type, SHAPES.interior);
    const pending = await createDocx(fromProseDoc(toProseDoc(source), source));
    const pendingXml = await documentXml(pending);
    const wrapperTag =
      type === "moveFrom"
        ? "moveFrom"
        : type === "moveTo"
          ? "moveTo"
          : type === "insertion"
            ? "ins"
            : "del";
    expect(pendingXml).toContain(`<w:${wrapperTag} `);
    expect(pendingXml).toContain('<w:br w:type="page" w:clear="all"/>');
    expect(pendingXml).toContain("<w:rPrChange ");
    expect(pendingXml).toContain('<w:rPrChange w:id="92"');
    expect(firstTrackedRun(await parseDocx(pending)).run.propertyChanges?.at(0)?.info.id).toBe(
      PROPERTY_REVISION.info.id,
    );

    const reviewer = await FolioDocxReviewer.fromBuffer(pending);
    expect(reviewer.getChanges()).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: OUTER_REVISION.id,
          type: type === "deletion" || type === "moveFrom" ? "deletion" : "insertion",
        }),
        expect.objectContaining({ id: PROPERTY_REVISION.info.id, type: "formatting" }),
      ]),
    );
    expect(reviewer.acceptChange(PROPERTY_REVISION.info.id)).toBe(true);
    const formattingResolvedXml = await documentXml(await reviewer.toBuffer());
    expect(formattingResolvedXml).toContain(`<w:${wrapperTag} `);
    expect(formattingResolvedXml).not.toContain("<w:rPrChange ");
    expect(pageBreakCount(formattingResolvedXml)).toBe(1);
  });

  test.each(WRAPPERS)(
    "accept/reject resolves stacked revisions and reaches both clean projections for %s",
    async (type) => {
      const source = trackedDocument(type, SHAPES.interior);
      const pending = await createDocx(fromProseDoc(toProseDoc(source), source));
      const keepOnAccept = type === "insertion" || type === "moveTo";

      for (const decision of ["accept", "reject"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(pending);
        expect(decision === "accept" ? reviewer.acceptAll() : reviewer.rejectAll()).toBe(2);
        const resolved = await reviewer.toBuffer();
        const reopened = await FolioDocxReviewer.fromBuffer(resolved);
        expect(reopened.getChanges()).toHaveLength(0);
        const reparsed = await parseDocx(resolved);
        const paragraph = reparsed.package.document.content.at(0);
        if (paragraph?.type !== "paragraph") {
          throw new Error("Expected resolved paragraph");
        }
        const keepsBreak = decision === "accept" ? keepOnAccept : !keepOnAccept;
        expect(cleanParagraphProjection(paragraph)).toEqual(keepsBreak ? ["A", "<page>", "B"] : []);
        expect(pageBreakCount(await documentXml(resolved))).toBe(keepsBreak ? 1 : 0);
        if (keepsBreak) {
          const run = paragraph.content.at(0);
          if (run?.type !== "run") {
            throw new Error("Expected resolved run");
          }
          expect(run.formatting).toMatchObject(
            decision === "accept" ? { underline: { style: "single" } } : { italic: true },
          );
        }
      }
    },
  );

  test.each(WRAPPERS)("reaches a save/reopen fixed point for a pending %s", async (type) => {
    const source = trackedDocument(type, SHAPES.multiple);
    const first = await createDocx(fromProseDoc(toProseDoc(source), source));
    const reopened = await parseDocx(first);
    const second = await createDocx(fromProseDoc(toProseDoc(reopened), reopened));

    expect(await documentXml(second)).toBe(await documentXml(first));
  });

  test("serializes the same owned page-break graph deterministically", async () => {
    const source = trackedDocument("insertion", SHAPES.multiple);
    const roundTripped = fromProseDoc(toProseDoc(source), source);

    expect(await documentXml(await createDocx(roundTripped))).toBe(
      await documentXml(await createDocx(roundTripped)),
    );
  });
});
