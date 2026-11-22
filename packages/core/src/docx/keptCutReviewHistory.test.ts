/**
 * A kept range marker's cut (`keptCut`) is review history: a save writes it
 * as `folio:keptCut`, a reopen reads it back, and resolving the reopened file
 * places the marker as resolving the open one does. The editor model does not
 * carry it.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";

import { escapeXmlAttribute, validateDocxPackage } from "@stll/docx-core";
import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { assertExactModel } from "../../../../test/exactModel";
import { propertyTestTimeout } from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, ParagraphContent } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseNumbering } from "./numberingParser";
import { parseParagraph } from "./paragraphParser";
import { parseDocx } from "./parser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { reviewProvenanceIssue } from "./reviewResolutionProvenance";
import { createDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CUT = { keptFrom: "2026-03-04T05:00:01Z" };

type LinkChildren = Extract<
  ParagraphContent,
  { type: "hyperlink" }
>["children"];

const link = (children: LinkChildren): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children,
});

const text = (value: string) =>
  ({
    type: "run",
    content: [{ type: "text", text: value }],
  }) satisfies ParagraphContent;

const keptCuts = (value: unknown): unknown[] => {
  if (Array.isArray(value)) return value.flatMap(keptCuts);
  if (typeof value !== "object" || value === null) return [];
  return Object.entries(value).flatMap(([key, field]) =>
    key === "keptCut" ? [field] : keptCuts(field),
  );
};

const reread = (paragraph: Paragraph): Paragraph => {
  const saved = serializeParagraph(paragraph);
  const root = parseXmlDocument(
    saved.replace(
      "<w:p",
      `<w:p xmlns:w="${W}" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`,
    ),
  ) as XmlElement | null;
  if (!root) throw new Error("The saved paragraph parses.");
  return parseParagraph(root, null, null, null, null, null);
};

describe("a kept range marker's cut", () => {
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [
      { type: "bookmarkStart", id: 1, name: "kept", keptCut: CUT },
      text("a"),
      { type: "bookmarkEnd", id: 1, keptCut: { ...CUT, piece: "before" } },
      { type: "commentRangeStart", id: 2, keptCut: CUT },
      {
        type: "moveToRangeStart",
        id: 3,
        name: "move3",
        author: "A",
        keptCut: CUT,
      },
      text("b"),
      { type: "moveToRangeEnd", id: 3, keptCut: { ...CUT, piece: "after" } },
      { type: "commentRangeEnd", id: 2, keptCut: { ...CUT, piece: "both" } },
      link([
        text("c"),
        { type: "bookmarkStart", id: 4, name: "inLink", keptCut: CUT },
        { type: "bookmarkEnd", id: 4, keptCut: CUT },
      ]),
    ],
  };

  test("is written as review history and read back on every marker", () => {
    const saved = serializeParagraph(paragraph);
    expect(saved.match(/folio:keptCut=/gu)).toHaveLength(8);
    const reopened = reread(paragraph);
    expect(reopened.content.map(({ type }) => type)).toEqual(
      paragraph.content.map(({ type }) => type),
    );
    expect(keptCuts(reopened)).toEqual(keptCuts(paragraph));
  });

  const readMarker = (value: unknown): Paragraph["content"] => {
    const encoded = escapeXmlAttribute(JSON.stringify({ version: 2, value }));
    const root = parseXmlDocument(
      `<w:p xmlns:w="${W}" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"><w:bookmarkEnd w:id="1" folio:keptCut="${encoded}"/></w:p>`,
    ) as XmlElement | null;
    if (!root) throw new Error("The paragraph parses.");
    return parseParagraph(root, null, null, null, null, null).content;
  };

  test.each([
    ["kept from no revision", { piece: "before" }],
    ["naming no side as its piece", { ...CUT, piece: "inside" }],
  ])("an unreadable cut (%s) is dropped and the marker kept", (_, value) => {
    expect(readMarker(value)).toEqual([{ type: "bookmarkEnd", id: 1 }]);
  });

  test.each([3, 0])(
    "a cut saved with a depth (%s) is read without it",
    (depth) => {
      const value = { depth, ...CUT, piece: "before" };
      expect(readMarker(value)).toEqual([
        { type: "bookmarkEnd", id: 1, keptCut: { ...CUT, piece: "before" } },
      ]);
      expect(
        reviewProvenanceIssue("keptCut", JSON.stringify({ version: 2, value })),
      ).toBeUndefined();
    },
  );

  test("a bookmark inside a link is read as one outside it", () => {
    const inLink: Paragraph = {
      type: "paragraph",
      paraId: "00000001",
      content: [
        link([
          {
            type: "bookmarkStart",
            id: 5,
            name: "copied",
            displacedByCustomXml: "next",
            copyOriginAncestors: [0],
          },
          text("d"),
          { type: "bookmarkEnd", id: 5, copyOriginAncestors: [0] },
        ]),
      ],
    };
    const childrenOf = ({ content }: Paragraph) =>
      content.flatMap((node) =>
        node.type === "hyperlink" ? node.children : [],
      );
    // The link's target comes from the package relationships, not read here.
    expect(childrenOf(reread(inLink))).toEqual(childrenOf(inLink));
  });

  test("the editor model does not carry it", () => {
    const input: Document = { package: { document: { content: [paragraph] } } };
    const back = fromProseDoc(toProseDoc(input), input, {
      stylesheetSource: { type: "package" },
    });
    expect(keptCuts(back)).toEqual([]);
  });
});

const documentWith = (content: Paragraph["content"]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { type: "paragraph", paraId: "00000001", content },
  ];
  return document;
};

const KEPT_FROM = "2026-03-04T05:00:02Z";

/**
 * A link piece in a pending deletion, a kept bookmark end, an insertion
 * waiting with the seam, the other piece: the seam's kept boundaries were
 * kept from a revision dated {@link KEPT_FROM}, the waiting insertion's side
 * the piece of their cut.
 */
const tracked = (): Document =>
  documentWith([
    {
      type: "deletion",
      info: { id: 20, author: "B", date: "2026-03-04T05:00:01Z" },
      content: [link([text("x")])],
      resolutionJoins: { before: 0, after: 2, remove: 0 },
    },
    {
      type: "bookmarkEnd",
      id: 1,
      keptCut: { keptFrom: KEPT_FROM, piece: "after" },
    },
    {
      type: "insertion",
      info: { id: 6, author: "A", date: "2026-03-04T05:00:03Z" },
      content: [text("y")],
      resolutionJoins: {
        before: 0,
        after: 0,
        remove: 0,
        deferredRemove: [{ depth: 2, blockers: [6] }],
      },
    },
    link([text("a")]),
  ]);

/**
 * The same document as folio review history version 1 wrote it: the seam's
 * kept date on the waiting insertion, no cut on the marker.
 */
const VERSION_1_PARAGRAPH = [
  '<w:p w14:paraId="00000001">',
  `<w:del w:id="20" w:author="B" w:date="2026-03-04T05:00:01Z" folio:resolutionJoins="${escapeXmlAttribute(JSON.stringify({ version: 1, value: { before: 0, after: 2, remove: 0 } }))}">`,
  '<w:hyperlink r:id="rId9"><w:r><w:delText>x</w:delText></w:r></w:hyperlink></w:del>',
  '<w:bookmarkEnd w:id="1"/>',
  `<w:ins w:id="6" w:author="A" w:date="2026-03-04T05:00:03Z" folio:resolutionJoins="${escapeXmlAttribute(JSON.stringify({ version: 1, value: { before: 0, after: 0, remove: 0, deferredRemove: [{ depth: 2, blockers: [6], keptFrom: KEPT_FROM }] } }))}">`,
  "<w:r><w:t>y</w:t></w:r></w:ins>",
  '<w:hyperlink r:id="rId9"><w:r><w:t>a</w:t></w:r></w:hyperlink>',
  "</w:p>",
].join("");

const paragraphOf = (document: Document) => {
  const first = document.package.document.content[0];
  return first?.type === "paragraph" ? first.content : [];
};

/** The paragraph's records, text, boundaries and cuts, without package-level ids. */
const shape = (nodes: readonly unknown[]): unknown[] =>
  nodes.map((node) => {
    if (typeof node !== "object" || node === null) return node;
    const record: Record<string, unknown> = Object.fromEntries(
      Object.entries(node),
    );
    const children = record["content"] ?? record["children"];
    return {
      type: record["type"],
      ...(typeof record["text"] === "string" ? { text: record["text"] } : {}),
      ...(typeof record["id"] === "number" ? { id: record["id"] } : {}),
      ...(record["keptCut"] === undefined ? {} : { cut: record["keptCut"] }),
      ...(Array.isArray(children) ? { children: shape(children) } : {}),
    };
  });

type Step = readonly [number, RevisionDecision];

const resolveOne = (document: Document, [id, decision]: Step): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [id],
    decision,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const saveAndReopen = async (document: Document): Promise<Document> => {
  const saved = await createDocx(document);
  expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({
    valid: true,
  });
  return parseDocx(saved, { preloadFonts: false });
};

const { ACCEPT, REJECT } = REVISION_DECISIONS;
const ORDERS: readonly (readonly [Step, Step])[] = [6, 20].flatMap((first) =>
  [ACCEPT, REJECT].flatMap((firstDecision) =>
    [ACCEPT, REJECT].map((secondDecision): readonly [Step, Step] => [
      [first, firstDecision],
      [first === 6 ? 20 : 6, secondDecision],
    ]),
  ),
);

describe("a review saved midway", () => {
  test("places kept markers as the open review does, in every order", async () => {
    let cutsSaved = 0;
    for (const [first, second] of ORDERS) {
      const midway = resolveOne(tracked(), first);
      const reopened = await saveAndReopen(midway);
      const marked = keptCuts(paragraphOf(midway));
      cutsSaved += marked.length;
      expect(keptCuts(paragraphOf(reopened))).toEqual(marked);
      assertExactModel(
        shape(paragraphOf(resolveOne(reopened, second))),
        shape(paragraphOf(resolveOne(midway, second))),
      );
    }
    // Some order leaves a marker standing in its cut across the save.
    expect(cutsSaved).toBeGreaterThan(0);
  });

  test.each(["legacy", "canonical"] as const)(
    "a file whose cut was saved with a depth opens, and saves unedited as it is (%s)",
    async (session) => {
      const zip = await ZipArchive.read(await createDocx(tracked()));
      const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
      const cut = { keptFrom: KEPT_FROM, piece: "after" };
      const encoded = (value: object) =>
        escapeXmlAttribute(JSON.stringify({ version: 2, value }));
      const older = xml.replace(encoded(cut), encoded({ depth: 2, ...cut }));
      expect(older).not.toBe(xml);
      zip.set("word/document.xml", older);
      const file = (await zip.generate({ compression: "store" })).buffer;
      const reopened = await parseDocx(file, { preloadFonts: false });
      expect(keptCuts(paragraphOf(reopened))).toEqual([cut]);
      const reviewer = await FolioDocxReviewer.fromBuffer(file, { session });
      const saved = await ZipArchive.read(await reviewer.toBuffer());
      expect(await saved.entry("word/document.xml")?.text()).toBe(older);
    },
  );

  test("a version 1 file reopens and resolves as the review it was saved from", async () => {
    const zip = await ZipArchive.read(await createDocx(tracked()));
    const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
    // The saved review declares the history namespace the fixture uses.
    expect(xml).toContain(`xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`);
    const body = /<w:body>(.*)<w:sectPr/su.exec(xml);
    if (body?.[1] === undefined) throw new Error("The body has a paragraph.");
    zip.set("word/document.xml", xml.replace(body[1], VERSION_1_PARAGRAPH));
    const reopened = await parseDocx(
      (await zip.generate({ compression: "store" })).buffer,
      { preloadFonts: false },
    );
    for (const steps of ORDERS) {
      let fromFile = reopened;
      let fromSource = tracked();
      for (const step of steps) {
        fromFile = resolveOne(fromFile, step);
        fromSource = resolveOne(fromSource, step);
      }
      assertExactModel(
        shape(paragraphOf(fromFile)),
        shape(paragraphOf(fromSource)),
      );
    }
  });
});

describe("a version 1 change naming its kept markers", () => {
  const R =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const LIST_NUMBER = [
    '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> LISTNUM </w:instrText></w:r>',
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>(a)</w:t></w:r>',
    '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:tab/></w:r>',
  ].join("");
  const NUMBERING = parseNumbering(
    [
      `<w:numbering xmlns:w="${W}">`,
      '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>',
      '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>',
      "</w:numbering>",
    ].join(""),
  );
  const read = (version: number, listNumber?: "listNumber"): Paragraph => {
    const joins = {
      before: 1,
      after: 0,
      remove: 0,
      keptFrom: KEPT_FROM,
      keptCount: 1,
    };
    const root = parseXmlDocument(
      [
        `<w:p xmlns:w="${W}" xmlns:r="${R}" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}">`,
        listNumber === undefined ? "" : LIST_NUMBER,
        '<w:bookmarkEnd w:id="1"/>',
        `<w:ins w:id="3" w:author="A" w:date="2026-03-04T05:00:03Z" folio:resolutionJoins="${escapeXmlAttribute(JSON.stringify({ version, value: joins }))}">`,
        '<w:hyperlink r:id="rId9"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:ins>',
        "</w:p>",
      ].join(""),
    ) as XmlElement | null;
    if (!root) throw new Error("The paragraph parses.");
    return parseParagraph(
      root,
      null,
      null,
      listNumber === undefined ? null : NUMBERING,
      null,
      null,
    );
  };

  test("marks them as standing in its cut in a paragraph whose marker shows a LISTNUM field", () => {
    const numbered = read(1, "listNumber");
    // The field and its tab fold into the marker.
    expect(numbered.listRendering?.foldedMarkerSuffix).toBe("(a)");
    expect(numbered.content.at(-2)).toEqual({
      type: "bookmarkEnd",
      id: 1,
      keptCut: { keptFrom: KEPT_FROM, piece: "after" },
    });
  });

  test("marks them as standing in its cut, its side the piece", () => {
    expect(read(1).content.at(0)).toEqual({
      type: "bookmarkEnd",
      id: 1,
      keptCut: { keptFrom: KEPT_FROM, piece: "after" },
    });
    // The model holds the joins without what version 1 said of the markers.
    expect(read(1).content.at(1)).toMatchObject({
      resolutionJoins: { before: 1, after: 0, remove: 0 },
    });
    expect(read(1).content.at(1)).not.toHaveProperty(
      "resolutionJoins.keptFrom",
    );
    // Version 2 marks kept markers itself: an unmarked one was not kept, and
    // joins naming markers are no version 2 joins.
    expect(read(2).content.at(0)).toEqual({ type: "bookmarkEnd", id: 1 });
    expect(read(2).content.at(1)).not.toHaveProperty("resolutionJoins");
  });

  test("accepting the change places them inside its piece", () => {
    const accepted = resolveOne(documentWith(read(1).content), [
      3,
      REVISION_DECISIONS.ACCEPT,
    ]);
    expect(shape(paragraphOf(accepted))).toEqual([
      {
        type: "hyperlink",
        children: [
          { type: "bookmarkEnd", id: 1 },
          { type: "run", children: [{ type: "text", text: "x" }] },
        ],
      },
    ]);
  });
});
