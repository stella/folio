/**
 * A canonical save writes `word/document.xml` by replaying the changed body
 * blocks onto the source part, and the selective save then routes the changed
 * paragraphs from that serialization back into the source part. The routing
 * reads the serialized side from the source part's scan and the regions the
 * replay replaced, instead of scanning and parsing the serialized part again.
 *
 *   1. For generated source parts (tables, content controls, text boxes,
 *      alternate content, id-less paragraphs, repeated ids, comments that
 *      spell paragraph and cell tags) and chains of edits, the routed part has
 *      exactly the bytes of the routing that scans and parses the serialized
 *      part whole, kept only as the oracle in
 *      __fixtures__/selectiveRoutingWholePartScan.ts. The replay's splices
 *      applied to its source give the part it wrote.
 *   2. Fixed cases reach each way the regions can or cannot stand in for the
 *      serialized part.
 *   3. Work per routed save: the source part is scanned and parsed whole once
 *      and the serialized part not at all, and a canonical selective save
 *      hands the router the replay.
 */

import { describe, expect, setDefaultTimeout, spyOn, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { buildPatchedDocumentXml as wholePartScanRouting } from "./__fixtures__/selectiveRoutingWholePartScan";
import { getDocumentSourceBaseline } from "./headerFooterVerbatim";
import { parseDocx } from "./parser";
import { createEmptyDocx } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";
import * as selectiveXmlPatch from "./selectiveXmlPatch";
import {
  buildPatchedDocumentXml,
  type SerializedPart,
} from "./selectiveXmlPatch";
import { serializeDocumentPart } from "./serializer/documentSerializer";
import { readRootNamespaceBindings } from "./serializer/partNamespaces";
import * as wordprocessingPrefixes from "./wordprocessingPrefixes";
import * as xmlParser from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(60_000));

const NAMESPACES = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  v: "urn:schemas-microsoft-com:vml",
} as const;

type ParagraphSpec = {
  paraId: string | undefined;
  text: string;
  /** A comment inside the paragraph that a literal tag scan reads as a tag. */
  trap: "none" | "cell" | "paragraph";
};
type BlockSpec =
  | ({ kind: "paragraph" } & ParagraphSpec)
  | { kind: "table"; cells: ParagraphSpec[][] }
  | { kind: "sdt"; paragraphs: ParagraphSpec[] }
  | { kind: "bookmark"; id: number }
  | { kind: "textBox"; paraId: string | undefined; inner: ParagraphSpec }
  | {
      kind: "alternate";
      paraId: string | undefined;
      choice: ParagraphSpec;
      fallback: ParagraphSpec;
    };
type PartSpec = {
  blocks: { gap: string; block: BlockSpec }[];
  sectionProperties: boolean;
};

const paragraphXml = (value: ParagraphSpec, inner = ""): string =>
  `<w:p${value.paraId === undefined ? "" : ` w14:paraId="${value.paraId}"`}>` +
  {
    none: "",
    cell: "<!-- <w:tc> -->",
    paragraph: "<!-- <w:p> -->",
  }[value.trap] +
  inner +
  `<w:r><w:t>${value.text}</w:t></w:r></w:p>`;

const textBoxRun = (inner: ParagraphSpec): string =>
  `<w:r><w:pict><v:shape style="width:100pt;height:40pt"><v:textbox><w:txbxContent>${paragraphXml(inner)}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;

const blockXml = (value: BlockSpec): string => {
  switch (value.kind) {
    case "paragraph":
      return paragraphXml(value);
    case "table":
      return (
        `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>` +
        `<w:gridCol w:w="2000"/>`.repeat(value.cells.length) +
        `</w:tblGrid><w:tr>${value.cells
          .map(
            (cell) =>
              `<w:tc>${cell.map((p) => paragraphXml(p)).join("")}</w:tc>`,
          )
          .join("")}</w:tr></w:tbl>`
      );
    case "sdt":
      return `<w:sdt><w:sdtPr><w:alias w:val="Clause"/></w:sdtPr><w:sdtContent>${value.paragraphs
        .map((p) => paragraphXml(p))
        .join("")}</w:sdtContent></w:sdt>`;
    case "bookmark":
      return `<w:bookmarkStart w:id="${value.id}" w:name="b${value.id}"/><w:bookmarkEnd w:id="${value.id}"/>`;
    case "textBox":
      return paragraphXml(
        { paraId: value.paraId, text: "box", trap: "none" },
        textBoxRun(value.inner),
      );
    case "alternate":
      return paragraphXml(
        { paraId: value.paraId, text: "alternate", trap: "none" },
        `<mc:AlternateContent><mc:Choice Requires="w14">${textBoxRun(value.choice)}</mc:Choice>` +
          `<mc:Fallback>${textBoxRun(value.fallback)}</mc:Fallback></mc:AlternateContent>`,
      );
    default: {
      const unreachable: never = value;
      return unreachable;
    }
  }
};

const writePart = (spec: PartSpec): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<w:document${Object.entries(NAMESPACES)
    .map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`)
    .join("")} mc:Ignorable="w14"><w:body>` +
  spec.blocks.map(({ gap, block }) => gap + blockXml(block)).join("") +
  (spec.sectionProperties
    ? `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`
    : "") +
  `</w:body></w:document>`;

const PARA_IDS = ["1A2B3C01", "1A2B3C02", "1A2B3C03", "1A2B3C04"] as const;
const TEXTS = ["same", "other", "third"] as const;
const GAPS = [
  "",
  "",
  "\n",
  "\n",
  " <!-- authored gap --> ",
  "<!-- <w:p> -->",
] as const;

const paragraphSpecArb: fc.Arbitrary<ParagraphSpec> = fc.record({
  paraId: fc.option(fc.constantFrom(...PARA_IDS), { nil: undefined }),
  text: fc.constantFrom(...TEXTS),
  trap: fc.constantFrom("none" as const, "none" as const, "cell", "paragraph"),
});
const paraIdArb = fc.option(fc.constantFrom(...PARA_IDS), { nil: undefined });
const blockSpecArb: fc.Arbitrary<BlockSpec> = fc.oneof(
  {
    weight: 5,
    arbitrary: paragraphSpecArb.map((value) => ({
      kind: "paragraph" as const,
      ...value,
    })),
  },
  {
    weight: 1,
    arbitrary: fc
      .array(fc.array(paragraphSpecArb, { minLength: 1, maxLength: 2 }), {
        minLength: 1,
        maxLength: 2,
      })
      .map((cells) => ({ kind: "table" as const, cells })),
  },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphSpecArb, { minLength: 1, maxLength: 2 })
      .map((paragraphs) => ({ kind: "sdt" as const, paragraphs })),
  },
  {
    weight: 1,
    arbitrary: fc
      .integer({ min: 1, max: 3 })
      .map((id) => ({ kind: "bookmark" as const, id })),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      kind: fc.constant("textBox" as const),
      paraId: paraIdArb,
      inner: paragraphSpecArb,
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      kind: fc.constant("alternate" as const),
      paraId: paraIdArb,
      choice: paragraphSpecArb,
      fallback: paragraphSpecArb,
    }),
  },
);
const partSpecArb: fc.Arbitrary<PartSpec> = fc.record({
  blocks: fc.array(
    fc.record({ gap: fc.constantFrom(...GAPS), block: blockSpecArb }),
    { minLength: 1, maxLength: 7 },
  ),
  sectionProperties: fc.boolean(),
});

type Edit =
  | { kind: "text"; at: number; text: string }
  | { kind: "insert"; at: number; paraId: string | undefined }
  | { kind: "delete"; at: number }
  | { kind: "move"; from: number; to: number }
  | { kind: "section"; pageWidth: number };

const editArb: fc.Arbitrary<Edit> = fc.oneof(
  {
    weight: 4,
    arbitrary: fc.record({
      kind: fc.constant("text" as const),
      at: fc.nat(8),
      text: fc.constantFrom(...TEXTS, "edited"),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      kind: fc.constant("insert" as const),
      at: fc.nat(8),
      paraId: fc.option(fc.constantFrom(...PARA_IDS, "1A2B3C09"), {
        nil: undefined,
      }),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      kind: fc.constant("delete" as const),
      at: fc.nat(8),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      kind: fc.constant("move" as const),
      from: fc.nat(8),
      to: fc.nat(8),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      kind: fc.constant("section" as const),
      pageWidth: fc.constantFrom(11906, 12240),
    }),
  },
);

const textRuns = (text: string): Paragraph["content"] => [
  { type: "run", content: [{ type: "text", text }] },
];

/** The paragraph an edit of `block` rewrites. */
const editedParagraph = (block: BlockContent): Paragraph | undefined => {
  switch (block.type) {
    case "paragraph":
      return block;
    case "table": {
      const first = block.rows.at(0)?.cells.at(0)?.content.at(0);
      return first?.type === "paragraph" ? first : undefined;
    }
    case "blockSdt":
    case "blockCustomXml": {
      const first = block.content.at(0);
      return first?.type === "paragraph" ? first : undefined;
    }
    default:
      return undefined;
  }
};

/** Every paragraph id the model holds, nested ones included. */
const modelParaIds = (content: readonly BlockContent[]): string[] =>
  content.flatMap((block): string[] => {
    switch (block.type) {
      case "paragraph":
        return block.paraId === undefined ? [] : [block.paraId];
      case "table":
        return block.rows.flatMap((row) =>
          row.cells.flatMap((cell) => modelParaIds(cell.content)),
        );
      case "blockSdt":
      case "blockCustomXml":
        return modelParaIds(block.content);
      default:
        return [];
    }
  });

/** Apply `edit`; answers the ids of the paragraphs it rewrote. */
const applyEdit = (document: Document, edit: Edit): string[] => {
  const body = document.package.document;
  const { content } = body;
  const index = (at: number) =>
    content.length === 0 ? 0 : at % content.length;
  switch (edit.kind) {
    case "text": {
      const block = content[index(edit.at)];
      const paragraph = block && editedParagraph(block);
      if (!paragraph) return [];
      paragraph.content = textRuns(edit.text);
      // The editor mints a key for a paragraph the source wrote without one.
      paragraph.paraId ??= `7E0000${(edit.at % 16).toString(16).toUpperCase()}0`;
      return [paragraph.paraId];
    }
    case "insert":
      content.splice(edit.at % (content.length + 1), 0, {
        type: "paragraph",
        ...(edit.paraId === undefined ? {} : { paraId: edit.paraId }),
        content: textRuns("inserted"),
      });
      return edit.paraId === undefined ? [] : [edit.paraId];
    case "delete":
      if (content.length > 0) content.splice(index(edit.at), 1);
      return [];
    case "move": {
      if (content.length === 0) return [];
      const [moved] = content.splice(index(edit.from), 1);
      if (moved) content.splice(edit.to % (content.length + 1), 0, moved);
      return moved ? modelParaIds([moved]) : [];
    }
    case "section":
      body.finalSectionProperties = {
        ...body.finalSectionProperties,
        pageWidth: edit.pageWidth,
      };
      return [];
    default: {
      const unreachable: never = edit;
      return unreachable;
    }
  }
};

const openPart = async (xml: string): Promise<Document> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set("word/document.xml", xml);
  return parseDocx((await zip.generate({ compression: "store" })).buffer);
};

const sourceOf = (document: Document): string => {
  const baseline = getDocumentSourceBaseline(document);
  if (baseline.type !== "captured")
    throw new Error(`Source baseline ${baseline.type}`);
  return baseline.xml;
};

const canonicalPart = (document: Document, source: string): SerializedPart =>
  serializeDocumentPart(document, readRootNamespaceBindings(source), {
    bodyAuthority: "canonical",
    xml: source,
  });

const applySplices = (
  xml: string,
  splices: readonly selectiveXmlPatch.XmlSplice[],
): string => {
  let written = "";
  let kept = 0;
  for (const { start, end, newXml } of splices) {
    expect(start).toBeGreaterThanOrEqual(kept);
    written += xml.slice(kept, start) + newXml;
    kept = end;
  }
  return written + xml.slice(kept);
};

/** Route `changed` both ways and check the bytes agree; answers the routed part. */
const expectSameRouting = (
  source: string,
  part: SerializedPart,
  changed: ReadonlySet<string>,
): string | null => {
  if (part.replay !== undefined) {
    expect(part.replay.source).toBe(source);
    expect(applySplices(part.replay.source, part.replay.splices)).toBe(
      part.xml,
    );
  }
  const routed = buildPatchedDocumentXml(source, part, new Set(changed));
  expect(routed).toBe(wholePartScanRouting(source, part.xml, new Set(changed)));
  return routed;
};

/** Whole-part scans and parses of any of `parts` while `run` runs. */
const countWholePartReads = <T>(
  parts: readonly string[],
  run: () => T,
): { value: T; scans: number; parses: number } => {
  const scans = spyOn(wordprocessingPrefixes, "resolveSelectiveScanPrefixes");
  const parses = spyOn(xmlParser, "parseXmlDocument");
  try {
    const value = run();
    const whole = (calls: readonly (readonly unknown[])[]) =>
      calls.filter(([input]) => parts.some((part) => input === part)).length;
    return {
      value,
      scans: whole(scans.mock.calls),
      parses: whole(parses.mock.calls),
    };
  } finally {
    scans.mockRestore();
    parses.mockRestore();
  }
};

describe("routing a replayed document part", () => {
  test("writes the bytes of the routing that scans the serialized part whole, over edit chains", async () => {
    await fc.assert(
      fc.asyncProperty(
        partSpecArb,
        fc.array(
          fc.record({
            edit: editArb,
            extra: fc.array(fc.nat(), { maxLength: 3 }),
          }),
          { minLength: 1, maxLength: 5 },
        ),
        async (spec, steps) => {
          const document = await openPart(writePart(spec));
          const source = sourceOf(document);
          const changed = new Set<string>();
          for (const { edit, extra } of steps) {
            for (const id of applyEdit(document, edit)) changed.add(id);
            // Ids the editor did not report: unedited, repeated, absent, minted.
            const candidates = [
              ...modelParaIds(document.package.document.content),
              ...PARA_IDS,
              "7FFFFFFF",
            ];
            for (const pick of extra) {
              const id = candidates[pick % candidates.length];
              if (id !== undefined) changed.add(id);
            }
            expectSameRouting(source, canonicalPart(document, source), changed);
          }
        },
      ),
      propertyConfig({ numRuns: 150 }),
    );
  });

  const plain = (text: string, paraId?: string): ParagraphSpec => ({
    paraId,
    text,
    trap: "none",
  });
  const paragraph = (
    text: string,
    paraId?: string,
    trap: ParagraphSpec["trap"] = "none",
  ): BlockSpec => ({ kind: "paragraph", paraId, text, trap });
  const part = (blocks: BlockSpec[], gap = "\n"): PartSpec => ({
    blocks: blocks.map((block) => ({ gap, block })),
    sectionProperties: true,
  });
  type FixedCase = {
    name: string;
    part: PartSpec;
    edits: readonly Edit[];
    /** Whether the routing reads the serialized side from the replay. */
    fromReplay: boolean;
    /** Whether the changed paragraphs route, where the case is about that. */
    routes?: boolean;
  };
  const FIXED_CASES: readonly FixedCase[] = [
    {
      name: "one paragraph's text",
      part: part([
        paragraph("same", "1A2B3C01"),
        paragraph("other", "1A2B3C02"),
        paragraph("third", "1A2B3C03"),
      ]),
      edits: [{ kind: "text", at: 1, text: "edited" }],
      fromReplay: true,
      routes: true,
    },
    {
      name: "two adjacent paragraphs",
      part: part(
        [
          paragraph("same", "1A2B3C01"),
          paragraph("other", "1A2B3C02"),
          paragraph("third", "1A2B3C03"),
        ],
        "",
      ),
      edits: [
        { kind: "text", at: 0, text: "edited" },
        { kind: "text", at: 1, text: "edited" },
      ],
      fromReplay: true,
      routes: true,
    },
    {
      name: "a paragraph the source wrote without an id",
      part: part([paragraph("same", "1A2B3C01"), paragraph("other")]),
      edits: [{ kind: "text", at: 1, text: "edited" }],
      fromReplay: true,
      routes: true,
    },
    {
      name: "a table cell",
      part: part([
        paragraph("same", "1A2B3C01"),
        {
          kind: "table",
          cells: [[plain("other", "1A2B3C02")], [plain("third", "1A2B3C03")]],
        },
      ]),
      edits: [{ kind: "text", at: 1, text: "edited" }],
      fromReplay: true,
      routes: true,
    },
    {
      name: "a paragraph after a text box",
      part: part([
        {
          kind: "textBox",
          paraId: "1A2B3C01",
          inner: plain("box", "1A2B3C02"),
        },
        paragraph("other", "1A2B3C03"),
      ]),
      edits: [{ kind: "text", at: 1, text: "edited" }],
      fromReplay: true,
      routes: true,
    },
    {
      name: "a paragraph whose own comment reads as a cell, and one after it",
      part: part([
        paragraph("same", "1A2B3C01", "cell"),
        paragraph("other", "1A2B3C02"),
      ]),
      edits: [
        { kind: "text", at: 0, text: "edited" },
        { kind: "text", at: 1, text: "edited" },
      ],
      fromReplay: false,
    },
    {
      name: "a paragraph whose own comment reads as an open paragraph",
      part: part([
        paragraph("same", "1A2B3C01", "paragraph"),
        paragraph("other", "1A2B3C02"),
      ]),
      edits: [{ kind: "text", at: 0, text: "edited" }],
      fromReplay: false,
    },
    {
      name: "a paragraph after a gap comment that reads as an open paragraph",
      part: part(
        [paragraph("same", "1A2B3C01"), paragraph("other", "1A2B3C02")],
        "<!-- <w:p> -->",
      ),
      edits: [{ kind: "text", at: 1, text: "edited" }],
      fromReplay: false,
    },
    {
      name: "an inserted paragraph",
      part: part([
        paragraph("same", "1A2B3C01"),
        paragraph("other", "1A2B3C02"),
      ]),
      edits: [{ kind: "insert", at: 1, paraId: "1A2B3C09" }],
      fromReplay: true,
    },
    {
      name: "section properties beside a paragraph",
      part: part([
        paragraph("same", "1A2B3C01"),
        paragraph("other", "1A2B3C02"),
      ]),
      edits: [
        { kind: "text", at: 0, text: "edited" },
        { kind: "section", pageWidth: 11906 },
      ],
      fromReplay: false,
    },
  ];

  for (const fixed of FIXED_CASES) {
    test(`routes ${fixed.name} with the same bytes`, async () => {
      const document = await openPart(writePart(fixed.part));
      const source = sourceOf(document);
      const changed = new Set<string>();
      for (const edit of fixed.edits) {
        for (const id of applyEdit(document, edit)) changed.add(id);
      }
      const serialized = canonicalPart(document, source);
      expect(serialized.replay !== undefined).toBe(
        fixed.name !== "section properties beside a paragraph",
      );
      const { value: routed, ...serializedReads } = countWholePartReads(
        [serialized.xml],
        () => expectSameRouting(source, serialized, changed),
      );
      if (fixed.routes !== undefined)
        expect(routed !== null).toBe(fixed.routes);
      // The oracle scans the serialized part (and parses it, when it gets
      // that far); the routing reads it again only when the replay cannot
      // stand in for it.
      const reads = fixed.fromReplay ? 1 : 2;
      expect(serializedReads.scans).toBe(reads);
      if (routed !== null) expect(serializedReads.parses).toBe(reads);
    });
  }
});

describe("work per routed save", () => {
  const PARAGRAPHS = 40;
  const sourcePart = writePart({
    blocks: Array.from({ length: PARAGRAPHS }, (_, index) => ({
      gap: "\n",
      block: {
        kind: "paragraph" as const,
        paraId: (0x10000000 + index).toString(16).toUpperCase(),
        text: `paragraph ${index}`,
        trap: "none" as const,
      },
    })),
    sectionProperties: true,
  });

  const edit = (document: Document, at: number, text: string): string => {
    const block = document.package.document.content[at];
    const paragraph = block && editedParagraph(block);
    if (!paragraph?.paraId) throw new Error(`Missing paragraph ${at}`);
    paragraph.content = textRuns(text);
    return paragraph.paraId;
  };

  test("the source part is read whole once and the replayed part not at all", async () => {
    const document = await openPart(sourcePart);
    const source = sourceOf(document);
    const changed = new Set([
      edit(document, 5, "edited"),
      edit(document, 30, "edited"),
    ]);
    const part = canonicalPart(document, source);
    expect(part.replay?.splices).toHaveLength(2);
    const reads = (serialized: string | SerializedPart) => {
      const {
        value: routed,
        scans,
        parses,
      } = countWholePartReads([source, part.xml], () =>
        buildPatchedDocumentXml(source, serialized, changed),
      );
      expect(routed).not.toBeNull();
      expect(routed).toBe(wholePartScanRouting(source, part.xml, changed));
      return { scans, parses };
    };
    // The same routing handed only the serialized text reads both parts.
    expect(reads(part.xml)).toEqual({ scans: 2, parses: 2 });
    expect(reads(part)).toEqual({ scans: 1, parses: 1 });
  });

  test("a canonical selective save hands the router the replay", async () => {
    const document = await openPart(sourcePart);
    const original = document.originalBuffer;
    if (!original) throw new Error("Missing original buffer");
    const changed = new Set([edit(document, 12, "edited")]);
    const routing = spyOn(selectiveXmlPatch, "buildPatchedDocumentXml");
    try {
      const saved = await attemptSelectiveSave(document, original, {
        changedParaIds: changed,
        structuralChange: false,
        hasUntrackedChanges: false,
        bodyAuthority: "canonical",
      });
      expect(saved).not.toBeNull();
      expect(routing).toHaveBeenCalledTimes(1);
      const serialized = routing.mock.calls[0]?.[1];
      expect(typeof serialized).toBe("object");
      expect(
        typeof serialized === "object" ? serialized.replay?.splices : undefined,
      ).toHaveLength(1);
    } finally {
      routing.mockRestore();
    }
  });
});
