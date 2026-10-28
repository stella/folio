import { expect, test } from "bun:test";

import { InvalidOoxmlNumericIdError } from "@stll/docx-core";
import { applyDocumentOp, inlineLeafSpans } from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { validateFolioDocumentModel } from "./modelValidation";
import { visitInlineContentSlots } from "./paragraphTraversal";
import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";

const AUTHORED = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';
const WORD_NAMESPACES = {
  transitional: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  strict: "http://purl.oclc.org/ooxml/wordprocessingml/main",
} as const;
const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
<w:p w14:paraId="00000001"><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:t>target</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
<w:p w14:paraId="00000002"><w:hyperlink w:anchor="target"><w:moveFromRangeStart w:id="2" w:name="move1" ${AUTHORED}/><w:moveFrom w:id="10" ${AUTHORED}><w:r><w:delText>moved</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="2"/></w:hyperlink></w:p>
<w:p w14:paraId="00000003"><w:ins w:id="20" ${AUTHORED}><w:hyperlink w:anchor="target"><w:moveToRangeStart w:id="3" w:name="move1" ${AUTHORED}/><w:moveTo w:id="11" ${AUTHORED}><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="3"/></w:hyperlink></w:ins></w:p>
<w:sectPr/></w:body></w:document>`;

const parseSource = async (source = DOCUMENT_XML) => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set("word/document.xml", source);
  return parseDocx((await zip.generate({ compression: "store" })).buffer, {
    preloadFonts: false,
  });
};

const rangeFacts = (document: Document) =>
  document.package.document.content.flatMap((paragraph) => {
    if (paragraph.type !== "paragraph") return [];
    const linked = new WeakSet<object>();
    visitInlineContentSlots(paragraph, ({ item }) => {
      if (item.type !== "hyperlink") return;
      for (const { node } of inlineLeafSpans(item.children)) linked.add(node);
    });
    return inlineLeafSpans(paragraph.content).flatMap(({ node, before }) =>
      node.type === "moveFromRangeStart" ||
      node.type === "moveFromRangeEnd" ||
      node.type === "moveToRangeStart" ||
      node.type === "moveToRangeEnd"
        ? [
            {
              type: node.type,
              id: node.id,
              paragraph: paragraph.paraId,
              offset: before.offset,
              linked: linked.has(node),
            },
          ]
        : [],
    );
  });

const hyperlinkChildTypes = (document: Document): string[][] => {
  const links: string[][] = [];
  for (const paragraph of document.package.document.content) {
    if (paragraph.type !== "paragraph") continue;
    visitInlineContentSlots(paragraph, ({ item }) => {
      if (item.type === "hyperlink") {
        links.push(item.children.map((child) => child.type));
      }
    });
  }
  return links;
};

const expectedFacts = (text: string) => [
  {
    type: "moveFromRangeStart",
    id: 2,
    paragraph: "00000002",
    offset: 0,
    linked: true,
  },
  {
    type: "moveFromRangeEnd",
    id: 2,
    paragraph: "00000002",
    offset: text.length,
    linked: true,
  },
  {
    type: "moveToRangeStart",
    id: 3,
    paragraph: "00000003",
    offset: 0,
    linked: true,
  },
  {
    type: "moveToRangeEnd",
    id: 3,
    paragraph: "00000003",
    offset: text.length,
    linked: true,
  },
];

const roundtrip = async (document: Document) =>
  parseDocx(await repackDocx(document, { updateModifiedDate: false }), {
    preloadFonts: false,
  });

test("rejecting an enclosing insertion retains a moved range's hyperlink-authored positions", async () => {
  const parsed = await parseSource();
  expect(validateFolioDocumentModel(parsed).valid).toBe(true);
  const result = applyDocumentOp(parsed, {
    type: "resolveRevision",
    story: "main",
    revisionIds: [20],
    decision: "reject",
  });
  if (result.isErr()) throw result.error;
  const rejected = result.value.document;
  expect(rangeFacts(rejected)).toEqual([
    ...expectedFacts("moved").slice(0, 2),
    {
      type: "moveToRangeStart",
      id: 3,
      paragraph: "00000003",
      offset: 0,
      linked: false,
    },
    {
      type: "moveToRangeEnd",
      id: 3,
      paragraph: "00000003",
      offset: 0,
      linked: false,
    },
  ]);
  const saved = await repackDocx(rejected, { updateModifiedDate: false });
  const savedZip = await ZipArchive.read(saved);
  const xml = await savedZip.entry("word/document.xml")?.text();
  expect(xml?.match(/<w:moveToRange(?:Start|End)\b/gu) ?? []).toEqual([
    "<w:moveToRangeStart",
    "<w:moveToRangeEnd",
  ]);
  expect(rangeFacts(await roundtrip(rejected))).toEqual(rangeFacts(rejected));
});

const SOURCE_CASES = Object.values(WORD_NAMESPACES).flatMap((namespace) =>
  ["w", "word"].flatMap((prefix) =>
    ["", "moved", "界😀é", "ab"].map((text) => ({ namespace, prefix, text })),
  ),
);

test.each(SOURCE_CASES)(
  "move boundaries retain order, position and link ownership for %j",
  async ({ namespace, prefix, text }) => {
    const source = DOCUMENT_XML.replace(WORD_NAMESPACES.transitional, namespace)
      .replaceAll("w:", `${prefix}:`)
      .replace("xmlns:w=", `xmlns:${prefix}=`)
      .replaceAll(">moved<", `>${text}<`);
    const parsed = await parseSource(source);
    const edited = fromProseDoc(toProseDoc(parsed), parsed, {
      stylesheetSource: { type: "package" },
    });
    const reopened = await roundtrip(edited);
    for (const document of [parsed, edited, reopened]) {
      expect(validateFolioDocumentModel(document).valid).toBe(true);
      expect(rangeFacts(document)).toEqual(expectedFacts(text));
    }
  },
);

test("hyperlink move starts retain their authored move-bookmark attributes", async () => {
  const parsed = await parseSource();
  const edited = fromProseDoc(toProseDoc(parsed), parsed, {
    stylesheetSource: { type: "package" },
  });
  for (const document of [parsed, edited, await roundtrip(edited)]) {
    const starts = document.package.document.content.flatMap((paragraph) =>
      paragraph.type === "paragraph"
        ? inlineLeafSpans(paragraph.content).flatMap(({ node }) =>
            node.type === "moveFromRangeStart" ||
            node.type === "moveToRangeStart"
              ? [node]
              : [],
          )
        : [],
    );
    expect(starts).toEqual([
      {
        type: "moveFromRangeStart",
        id: 2,
        name: "move1",
        author: "Reviewer",
        date: "2026-01-01T00:00:00Z",
      },
      {
        type: "moveToRangeStart",
        id: 3,
        name: "move1",
        author: "Reviewer",
        date: "2026-01-01T00:00:00Z",
      },
    ]);
  }
});

test("empty move ranges stay in their indexed hyperlink beside same-href links", async () => {
  const source = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
<w:p w14:paraId="00000001"><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:t>target</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
<w:p w14:paraId="00000002"><w:hyperlink w:anchor="target"><w:moveFromRangeStart w:id="4" w:name="emptyFrom" ${AUTHORED}/><w:moveFromRangeEnd w:id="4" ${AUTHORED}/><w:r><w:t>untracked</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="target"><w:r><w:t>neighbor</w:t></w:r></w:hyperlink></w:p>
<w:p w14:paraId="00000003"><w:ins w:id="20" ${AUTHORED}><w:hyperlink w:anchor="target"><w:moveToRangeStart w:id="5" w:name="emptyTo" ${AUTHORED}/><w:moveToRangeEnd w:id="5" ${AUTHORED}/></w:hyperlink></w:ins><w:hyperlink w:anchor="target"><w:r><w:t>trackedNeighbor</w:t></w:r></w:hyperlink></w:p>
<w:sectPr/></w:body></w:document>`;
  const parsed = await parseSource(source);
  const edited = fromProseDoc(toProseDoc(parsed), parsed, {
    stylesheetSource: { type: "package" },
  });
  const reopened = await roundtrip(edited);
  const expectedRangeFacts = [
    {
      type: "moveFromRangeStart",
      id: 4,
      paragraph: "00000002",
      offset: 0,
      linked: true,
    },
    {
      type: "moveFromRangeEnd",
      id: 4,
      paragraph: "00000002",
      offset: 0,
      linked: true,
    },
    {
      type: "moveToRangeStart",
      id: 5,
      paragraph: "00000003",
      offset: 0,
      linked: true,
    },
    {
      type: "moveToRangeEnd",
      id: 5,
      paragraph: "00000003",
      offset: 0,
      linked: true,
    },
  ];
  for (const document of [parsed, edited, reopened]) {
    expect(validateFolioDocumentModel(document).valid).toBe(true);
    expect(rangeFacts(document)).toEqual(expectedRangeFacts);
    expect(hyperlinkChildTypes(document)).toEqual([
      ["moveFromRangeStart", "moveFromRangeEnd", "run"],
      ["run"],
      ["moveToRangeStart", "moveToRangeEnd"],
      ["run"],
    ]);
  }
});

test("an invalid hyperlink move marker retains its captured writer error", async () => {
  const parsed = await parseSource(
    DOCUMENT_XML.replace(
      '<w:moveToRangeEnd w:id="3"/>',
      '<w:moveToRangeEnd w:id="not-an-id"/>',
    ),
  );
  const captured: string[] = [];
  for (const paragraph of parsed.package.document.content) {
    if (paragraph.type !== "paragraph") continue;
    visitInlineContentSlots(paragraph, ({ item }) => {
      if (
        item.type === "preservedInline" &&
        item.xml.includes('w:id="not-an-id"')
      )
        captured.push(item.xml);
    });
  }
  expect(captured).toHaveLength(1);
  await expect(
    repackDocx(parsed, { updateModifiedDate: false }),
  ).rejects.toBeInstanceOf(InvalidOoxmlNumericIdError);
});

test("unsupported and unknown-namespace hyperlink markup stays captured beside supported moves", async () => {
  const parsed = await parseSource(
    DOCUMENT_XML.replace(
      '<w:moveToRangeEnd w:id="3"/>',
      '<x:moveToRangeEnd xmlns:x="urn:opaque" w:id="99"/><w:proofErr w:type="spellStart"/><w:proofErr w:type="spellEnd"/><w:moveToRangeEnd w:id="3"/>',
    ),
  );
  const saved = await repackDocx(parsed, { updateModifiedDate: false });
  const zip = await ZipArchive.read(saved);
  const xml = await zip.entry("word/document.xml")?.text();
  expect(xml).toContain("<x:moveToRangeEnd");
  expect(xml).toContain("urn:opaque");
  expect(xml).toContain('<w:proofErr w:type="spellStart"/>');
  expect(xml).toContain('<w:proofErr w:type="spellEnd"/>');
  expect(rangeFacts(await roundtrip(parsed))).toEqual(expectedFacts("moved"));
});

const moveMarkers = (document: Document, ends = false) =>
  document.package.document.content.flatMap((paragraph) =>
    paragraph.type === "paragraph"
      ? inlineLeafSpans(paragraph.content).flatMap(({ node }) =>
          node.type === "moveFromRangeEnd" ||
          node.type === "moveToRangeEnd" ||
          (!ends &&
            (node.type === "moveFromRangeStart" ||
              node.type === "moveToRangeStart"))
            ? [node]
            : [],
        )
      : [],
  );

test("hyperlink move ends retain their copy origin through the parser and the editor", async () => {
  const parsed = await parseSource();
  for (const end of moveMarkers(parsed, true))
    Object.assign(end, { copyOriginAncestors: [0] });
  const saved = await roundtrip(parsed);
  const edited = fromProseDoc(toProseDoc(saved), saved, {
    stylesheetSource: { type: "package" },
  });
  for (const document of [saved, edited, await roundtrip(edited)])
    expect(
      moveMarkers(document, true).map(({ type, copyOriginAncestors }) => ({
        type,
        copyOriginAncestors,
      })),
    ).toEqual([
      { type: "moveFromRangeEnd", copyOriginAncestors: [0] },
      { type: "moveToRangeEnd", copyOriginAncestors: [0] },
    ]);
});

test("each save of a hyperlink move marker is a record of its own", async () => {
  const parsed = await parseSource();
  const doc = toProseDoc(parsed);
  const save = () =>
    fromProseDoc(doc, parsed, { stylesheetSource: { type: "package" } });
  const first = moveMarkers(save());
  expect(first).toHaveLength(4);
  for (const marker of first) Object.assign(marker, { id: 99 });
  const second = moveMarkers(save());
  for (const marker of second) expect(first.includes(marker)).toBe(false);
  expect(second.map(({ id }) => id)).toEqual([2, 2, 3, 3]);
});
