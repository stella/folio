import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import { describe, expect, test } from "bun:test";

import {
  type DocxProjectionFactSet,
  type DocxProjectionOutlineLevelFact,
  type DocxProjectionParagraph,
  type DocxProjectionWire,
  initializeDocxProjection,
  projectCompressedDocx,
} from "../packages/docx-core/src/projection";
import { parseDocx } from "@stll/folio-core/docx/parser";
import type { BlockContent } from "@stll/folio-core/types/document";

import { DEFAULT_INVARIANT_BUDGET_MS } from "./lib/corpus-invariants/contract";
import {
  kernelDifferentialFailures,
  runKernelDifferentialInvariant,
  type TypeScriptDocumentFacts,
  typeScriptDocumentFacts,
} from "./lib/corpus-invariants/kernel-differential";

const UNKNOWN_FACTS = ["known", []] as const satisfies DocxProjectionFactSet<never>;

type ProjectionOptions = {
  paragraphs: readonly DocxProjectionParagraph[];
  outlineLevels?: DocxProjectionFactSet<DocxProjectionOutlineLevelFact>;
};

const projection = ({
  paragraphs,
  outlineLevels = ["known", []],
}: ProjectionOptions): DocxProjectionWire => [
  5,
  paragraphs,
  [UNKNOWN_FACTS, UNKNOWN_FACTS, UNKNOWN_FACTS, UNKNOWN_FACTS, outlineLevels],
  ["complete"],
  ["complete"],
];

type KernelParagraphOptions = {
  ordinal: number;
  styleId?: string | null;
  table?: readonly [tableId: string, row: number, column: number];
};

const kernelParagraph = ({
  ordinal,
  styleId = null,
  table,
}: KernelParagraphOptions): DocxProjectionParagraph => [
  ordinal,
  "",
  null,
  [],
  table === undefined ? [] : ["table", table[0], table[1], table[2]],
  styleId,
  null,
];

const typeScriptParagraph = (
  styleId: string | null = null,
  outlineLevel: number | null = null,
): TypeScriptDocumentFacts["paragraphs"][number] => ({ styleId, outlineLevel });

const messages = (wire: DocxProjectionWire, facts: TypeScriptDocumentFacts): string[] =>
  kernelDifferentialFailures(wire, facts).map(({ message }) => message);

describe("typeScriptDocumentFacts", () => {
  test("reads paragraphs in document order, cell paragraphs among body paragraphs", () => {
    const blocks = [
      { type: "paragraph", content: [], formatting: { styleId: "Before" } },
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [{ type: "paragraph", content: [], formatting: { styleId: "Inside" } }],
              },
            ],
          },
        ],
      },
      { type: "paragraph", content: [], formatting: { styleId: "After" } },
    ] satisfies BlockContent[];

    expect(typeScriptDocumentFacts(blocks).paragraphs.map(({ styleId }) => styleId)).toEqual([
      "Before",
      "Inside",
      "After",
    ]);
  });

  test("descends into a block content control, which the kernel treats as transparent", () => {
    const blocks = [
      {
        type: "blockSdt",
        properties: { sdtType: "richText" },
        content: [{ type: "paragraph", content: [], formatting: { styleId: "Controlled" } }],
      },
    ] satisfies BlockContent[];

    expect(typeScriptDocumentFacts(blocks).paragraphs).toEqual([typeScriptParagraph("Controlled")]);
  });

  test("numbers a nested table after its parent, matching the kernel's w:tbl counter", () => {
    const inner = {
      type: "table",
      rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [] }] }],
    } satisfies BlockContent;
    const blocks = [
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              { type: "tableCell", content: [inner] },
              { type: "tableCell", content: [] },
            ],
          },
        ],
      },
    ] satisfies BlockContent[];

    expect(typeScriptDocumentFacts(blocks).tables).toEqual([
      { cellsPerRow: [2] },
      { cellsPerRow: [1] },
    ]);
  });

  test("reports an absent style id and outline level as none, not as zero", () => {
    const blocks = [{ type: "paragraph", content: [] }] satisfies BlockContent[];

    expect(typeScriptDocumentFacts(blocks).paragraphs).toEqual([typeScriptParagraph(null, null)]);
  });
});

describe("kernelDifferentialFailures: paragraph counts", () => {
  test("claims nothing when the two walks disagree on how many paragraphs exist", () => {
    const wire = projection({ paragraphs: [kernelParagraph({ ordinal: 0, styleId: "A" })] });

    expect(
      messages(wire, {
        paragraphs: [typeScriptParagraph("B"), typeScriptParagraph("C")],
        tables: [],
      }),
    ).toEqual([]);
  });
});

describe("kernelDifferentialFailures: style ids", () => {
  test("says nothing when both read the same w:pStyle", () => {
    const wire = projection({ paragraphs: [kernelParagraph({ ordinal: 0, styleId: "Heading1" })] });

    expect(messages(wire, { paragraphs: [typeScriptParagraph("Heading1")], tables: [] })).toEqual(
      [],
    );
  });

  test("classifies the kernel reading no style where TypeScript read one", () => {
    const wire = projection({ paragraphs: [kernelParagraph({ ordinal: 0 })] });

    expect(messages(wire, { paragraphs: [typeScriptParagraph("Heading1")], tables: [] })).toEqual([
      "per-paragraph style id: the kernel reports none where TypeScript reports one",
    ]);
  });

  test("classifies the opposite direction separately", () => {
    const wire = projection({ paragraphs: [kernelParagraph({ ordinal: 0, styleId: "Heading1" })] });

    expect(messages(wire, { paragraphs: [typeScriptParagraph(null)], tables: [] })).toEqual([
      "per-paragraph style id: TypeScript reports none where the kernel reports one",
    ]);
  });

  test("reports a plain disagreement without naming either style", () => {
    const wire = projection({ paragraphs: [kernelParagraph({ ordinal: 0, styleId: "Heading1" })] });

    expect(messages(wire, { paragraphs: [typeScriptParagraph("Heading2")], tables: [] })).toEqual([
      "per-paragraph style id disagrees",
    ]);
  });

  test("reports one defect once, however many paragraphs carry it", () => {
    const wire = projection({
      paragraphs: [
        kernelParagraph({ ordinal: 0, styleId: "A" }),
        kernelParagraph({ ordinal: 1, styleId: "B" }),
        kernelParagraph({ ordinal: 2, styleId: "C" }),
      ],
    });

    expect(
      messages(wire, {
        paragraphs: [typeScriptParagraph("X"), typeScriptParagraph("Y"), typeScriptParagraph("Z")],
        tables: [],
      }),
    ).toEqual(["per-paragraph style id disagrees"]);
  });
});

describe("kernelDifferentialFailures: outline levels", () => {
  test("says nothing when the kernel resolves to the direct value folio holds", () => {
    const wire = projection({
      paragraphs: [kernelParagraph({ ordinal: 0 })],
      outlineLevels: ["known", [[0, 3]]],
    });

    expect(messages(wire, { paragraphs: [typeScriptParagraph(null, 3)], tables: [] })).toEqual([]);
  });

  test("reports a direct level the kernel resolved to something else", () => {
    const wire = projection({
      paragraphs: [kernelParagraph({ ordinal: 0 })],
      outlineLevels: ["known", [[0, 1]]],
    });

    expect(messages(wire, { paragraphs: [typeScriptParagraph(null, 3)], tables: [] })).toEqual([
      "direct outline level disagrees",
    ]);
  });

  test("reports a direct level the kernel gave no fact for", () => {
    const wire = projection({ paragraphs: [kernelParagraph({ ordinal: 0 })] });

    expect(messages(wire, { paragraphs: [typeScriptParagraph(null, 3)], tables: [] })).toEqual([
      "direct outline level: the kernel reports none where TypeScript reports one",
    ]);
  });

  test("accepts a level the kernel inherited and folio does not model", () => {
    const wire = projection({
      paragraphs: [kernelParagraph({ ordinal: 0 })],
      outlineLevels: ["known", [[0, 2]]],
    });

    expect(messages(wire, { paragraphs: [typeScriptParagraph(null, null)], tables: [] })).toEqual(
      [],
    );
  });

  test("treats an unknown family as the kernel declining to claim completeness", () => {
    const wire = projection({
      paragraphs: [kernelParagraph({ ordinal: 0 })],
      outlineLevels: ["unknown", "unsupported-styles"],
    });

    expect(messages(wire, { paragraphs: [typeScriptParagraph(null, 3)], tables: [] })).toEqual([]);
  });

  test("compares an empty known family, which is authoritative negative evidence", () => {
    const wire = projection({
      paragraphs: [kernelParagraph({ ordinal: 0 })],
      outlineLevels: ["known", []],
    });

    expect(messages(wire, { paragraphs: [typeScriptParagraph(null, 0)], tables: [] })).toEqual([
      "direct outline level: the kernel reports none where TypeScript reports one",
    ]);
  });
});

describe("kernelDifferentialFailures: table structure", () => {
  const inTable = (table: readonly [string, number, number]): DocxProjectionWire =>
    projection({ paragraphs: [kernelParagraph({ ordinal: 0, table })] });

  test("says nothing when the coordinate lies inside the modelled table", () => {
    expect(
      messages(inTable(["table-0", 1, 2]), {
        paragraphs: [typeScriptParagraph()],
        tables: [{ cellsPerRow: [1, 3] }],
      }),
    ).toEqual([]);
  });

  test("reports a table the TypeScript model does not have", () => {
    expect(
      messages(inTable(["table-1", 0, 0]), {
        paragraphs: [typeScriptParagraph()],
        tables: [{ cellsPerRow: [1] }],
      }),
    ).toEqual(["table structure: the kernel reports a table the TypeScript model lacks"]);
  });

  test("reports a row the TypeScript model does not have", () => {
    expect(
      messages(inTable(["table-0", 2, 0]), {
        paragraphs: [typeScriptParagraph()],
        tables: [{ cellsPerRow: [1, 1] }],
      }),
    ).toEqual(["table structure: the kernel reports a row the TypeScript model lacks"]);
  });

  test("reports a cell the TypeScript model does not have", () => {
    expect(
      messages(inTable(["table-0", 0, 4]), {
        paragraphs: [typeScriptParagraph()],
        tables: [{ cellsPerRow: [2] }],
      }),
    ).toEqual(["table structure: the kernel reports a cell the TypeScript model lacks"]);
  });

  test("accepts a row folio holds and the kernel never observed a paragraph in", () => {
    expect(
      messages(inTable(["table-0", 0, 0]), {
        paragraphs: [typeScriptParagraph()],
        tables: [{ cellsPerRow: [1, 1, 1] }],
      }),
    ).toEqual([]);
  });

  test("fails loudly rather than silently skipping an unrecognised table identifier", () => {
    expect(
      messages(inTable(["tbl_0", 0, 0]), {
        paragraphs: [typeScriptParagraph()],
        tables: [{ cellsPerRow: [1] }],
      }),
    ).toEqual(["kernel table identifier is not the projected shape"]);
  });
});

const packageFromDocumentXml = async (documentXml: string): Promise<Uint8Array> => {
  const archive = new JSZip();
  archive.file("word/document.xml", documentXml);
  return archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });
};

const runOver = async (documentXml: string) => {
  const bytes = await packageFromDocumentXml(documentXml);
  const buffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  return runKernelDifferentialInvariant({
    bytes,
    buffer,
    parsed: await parseDocx(buffer, { preloadFonts: false }),
    documentPart: "word/document.xml",
    budgetMs: DEFAULT_INVARIANT_BUDGET_MS,
  });
};

describe("runKernelDifferentialInvariant", () => {
  /**
   * Without this the passing runs below prove nothing: the comparison declines
   * to claim anything when the paragraph counts differ, so a walk that saw the
   * wrong paragraphs would also report zero failures.
   */
  test("walks the same paragraphs the kernel projects, cell paragraphs included", async () => {
    const bytes = await packageFromDocumentXml(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inside</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>`,
    );
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;
    await initializeDocxProjection({
      wasm: await readFile(
        new URL("../packages/docx-core/src/generated/docx_kernel_bg.wasm", import.meta.url),
      ),
    });

    const projected = await projectCompressedDocx(bytes);
    const parsed = await parseDocx(buffer, { preloadFonts: false });
    const facts = typeScriptDocumentFacts(parsed.package.document.content);

    expect(projected[1]).toHaveLength(3);
    expect(facts.paragraphs).toHaveLength(3);
    expect(facts.tables).toEqual([{ cellsPerRow: [1] }]);
  });

  test("finds nothing to report in a two-paragraph package carrying a style id", async () => {
    const { failures, timings } = await runOver(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Marked"/><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p></w:body></w:document>`,
    );

    expect(failures).toEqual([]);
    expect(Object.keys(timings).sort()).toEqual(["compare", "kernel-project"]);
  });

  test("finds nothing to report in a one-row one-cell table", async () => {
    const { failures } = await runOver(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inside</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`,
    );

    expect(failures).toEqual([]);
  });

  test("reports the kernel refusing a package the TypeScript parser accepted", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const { failures } = await runKernelDifferentialInvariant({
      bytes,
      buffer: bytes.buffer as ArrayBuffer,
      parsed: { package: { document: { content: [] } } },
      documentPart: "word/document.xml",
      budgetMs: DEFAULT_INVARIANT_BUDGET_MS,
    });

    expect(failures).toHaveLength(1);
    expect(failures.at(0)?.invariant).toBe("kernel-differential");
  });
});
