import { expect, test } from "bun:test";
import JSZip from "jszip";

import { InvalidOoxmlNumericIdError } from "@stll/docx-core";
import { applyDocumentOp, inlineLeafSpans } from "@stll/docx-core/ops";

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
  const zip = await JSZip.loadAsync(await createDocx(createEmptyDocument()));
  zip.file("word/document.xml", source);
  return parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
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
  const savedZip = await JSZip.loadAsync(saved);
  const xml = await savedZip.file("word/document.xml")?.async("text");
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
  const zip = await JSZip.loadAsync(saved);
  const xml = await zip.file("word/document.xml")?.async("text");
  expect(xml).toContain("<x:moveToRangeEnd");
  expect(xml).toContain("urn:opaque");
  expect(xml).toContain('<w:proofErr w:type="spellStart"/>');
  expect(xml).toContain('<w:proofErr w:type="spellEnd"/>');
  expect(rangeFacts(await roundtrip(parsed))).toEqual(expectedFacts("moved"));
});
