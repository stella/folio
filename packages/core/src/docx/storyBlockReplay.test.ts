import { expect, test } from "bun:test";
import type { BlockContent } from "../types/document";
import { parseHeader } from "./headerFooterParser";
import { buildStoryBlockReplay } from "./storyBlockReplay";
import { getChildElements, getNamespaceUri } from "./xmlParser";
import { parseStreamingXmlWithSourceRanges } from "./streamingXmlParser";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const block = (paraId: string | undefined, text: string) => {
  const value = {
    type: "paragraph",
    ...(paraId === undefined ? {} : { paraId }),
    content: [{ type: "run", content: [{ type: "text", text }] }],
  } satisfies BlockContent;
  const firstRun = value.content.at(0);
  if (firstRun) {
    const firstText = firstRun.content.at(0);
    if (firstText) Object.freeze(firstText);
    Object.freeze(firstRun.content);
    Object.freeze(firstRun);
  }
  Object.freeze(value.content);
  return Object.freeze(value);
};
const story = (body: string) => `<q:hdr xmlns:q="${WORD}">${body}</q:hdr>`;
const emitted = (body: string) => `<w:hdr xmlns:w="${WORD}">${body}</w:hdr>`;
const sourceParagraph = (text: string) => `<q:p><q:r><q:t>${text}</q:t></q:r></q:p>`;
const emittedParagraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

test("moves paragraphs with identical text by stable identity without reusing a source range", () => {
  const first = block("11111111", "same");
  const second = block("22222222", "same");
  const firstXml = `<q:p xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml" p14:paraId="11111111">\n<q:r><q:t>same</q:t></q:r></q:p>`;
  const secondXml = `<q:p xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml" p14:paraId="22222222"><q:r>\t<q:t>same</q:t></q:r></q:p>`;
  const sourceXml = story(`\n${firstXml}\n<!-- between -->\t${secondXml}\n\n`);
  expect(
    buildStoryBlockReplay({
      sourceXml,
      baselineContent: Object.freeze([first, second]),
      currentContent: Object.freeze([second, first]),
      serializedXml: emitted(emittedParagraph("same") + emittedParagraph("same")),
    }),
  ).toBe(story(`\n${secondXml}\n<!-- between -->\t${firstXml}\n\n`));
});

test("identical id-less blocks consume distinct original ranges once", () => {
  const value = block(undefined, "same");
  const firstXml = `<q:p>\n<q:r><q:t>same</q:t></q:r></q:p>`;
  const secondXml = `<q:p><q:r>\t<q:t>same</q:t></q:r></q:p>`;
  const sourceXml = story(firstXml + " " + secondXml);
  expect(
    buildStoryBlockReplay({
      sourceXml,
      baselineContent: [value, value],
      currentContent: [value, value, value],
      serializedXml: emitted(emittedParagraph("same").repeat(3)),
    }),
  ).toBe(
    story(firstXml + " " + secondXml + `<w:p xmlns:w="${WORD}"><w:r><w:t>same</w:t></w:r></w:p>`),
  );
});

test("insertions and deletions reflect the current sequence and retain all authored gaps", () => {
  const first = block("11111111", "first");
  const second = block("22222222", "second");
  const third = block("33333333", "third");
  const firstXml = sourceParagraph("first");
  const secondXml = sourceParagraph("second");
  const thirdXml = `<w:p xmlns:w="${WORD}"><w:r><w:t>third</w:t></w:r></w:p>`;
  const sourceXml = `<?xml version='1.0'?>\n${story(`\n${firstXml}\n<!-- gap -->\t${secondXml}\n\n`)}\n`;
  const output = buildStoryBlockReplay({
    sourceXml,
    baselineContent: [first, second],
    currentContent: [third, first, third],
    serializedXml: emitted(
      emittedParagraph("third") + emittedParagraph("first") + emittedParagraph("third"),
    ),
  });
  expect(output).toBe(
    `<?xml version='1.0'?>\n${story(`\n${thirdXml}\n<!-- gap -->\t${firstXml}${thirdXml}\n\n`)}\n`,
  );
  expect(
    buildStoryBlockReplay({
      sourceXml,
      baselineContent: [first, second],
      currentContent: [],
      serializedXml: emitted(""),
    }),
  ).toBe(`<?xml version='1.0'?>\n${story("\n\n<!-- gap -->\t\n\n")}\n`);
});

test("empty story roots accept insertion while preserving root attributes and trailing whitespace", () => {
  const value = block("11111111", "new");
  const fragment = `<w:p xmlns:w="${WORD}"><w:r><w:t>new</w:t></w:r></w:p>`;
  expect(
    buildStoryBlockReplay({
      sourceXml: story("\n\t"),
      baselineContent: [],
      currentContent: [value],
      serializedXml: emitted(emittedParagraph("new")),
    }),
  ).toBe(story(`\n\t${fragment}`));
  expect(
    buildStoryBlockReplay({
      sourceXml: `<q:hdr xmlns:q='${WORD}' />\n`,
      baselineContent: [],
      currentContent: [value],
      serializedXml: emitted(emittedParagraph("new")),
    }),
  ).toBe(`<q:hdr xmlns:q='${WORD}' >${fragment}</q:hdr>\n`);
});

test("changed fragments bind generated aliases locally and map namespace declarations to Strict", () => {
  const old = block("11111111", "old");
  const current = block("11111111", REL);
  const sourceXml = `<q:hdr xmlns:q='${STRICT_WORD}' xmlns:w="urn:foreign"><q:p><q:r><q:t>old</q:t></q:r></q:p>\n</q:hdr>`;
  const serializedXml = `<z:hdr xmlns:z="${WORD}" xmlns:r="${REL}"><z:p label=" xmlns:fake='${WORD}'"><z:hyperlink r:id="rId1"><z:r xmlns:local="${WORD}"><z:t>${REL}</z:t></z:r></z:hyperlink></z:p></z:hdr>`;
  const output = buildStoryBlockReplay({
    sourceXml,
    baselineContent: [old],
    currentContent: [current],
    serializedXml,
  });
  expect(output).toContain(`xmlns:z="${STRICT_WORD}"`);
  expect(output).toContain(`xmlns:r="${STRICT_REL}"`);
  expect(output).toContain(`xmlns:local="${STRICT_WORD}"`);
  expect(output).toContain(`<z:t>${REL}</z:t>`);
  expect(output?.startsWith(`<q:hdr xmlns:q='${STRICT_WORD}' xmlns:w="urn:foreign">`)).toBe(true);
  const parsed = parseStreamingXmlWithSourceRanges(output ?? "");
  expect(parsed.status).toBe("parsed");
  if (parsed.status !== "parsed") return;
  const root = getChildElements(parsed.value).at(0);
  const paragraph = getChildElements(root).at(0);
  const hyperlink = getChildElements(paragraph).at(0);
  expect(paragraph?.attributes?.["label"]).toBe(` xmlns:fake='${WORD}'`);
  expect(getNamespaceUri(paragraph ?? {})).toBe(STRICT_WORD);
  expect(getNamespaceUri(hyperlink ?? {})).toBe(STRICT_WORD);
});

test("unmapped elements, count mismatch and malformed XML fall back safely", () => {
  const value = block("11111111", "same");
  for (const sourceXml of [
    story(sourceParagraph("same") + "<q:proofErr/>"),
    story("<vendor:x xmlns:vendor='urn:vendor'/>"),
    story(""),
    "<q:hdr>",
  ]) {
    expect(
      buildStoryBlockReplay({
        sourceXml,
        baselineContent: [value],
        currentContent: [value],
        serializedXml: emitted(emittedParagraph("same")),
      }),
    ).toBeNull();
  }
  expect(
    buildStoryBlockReplay({
      sourceXml: story(sourceParagraph("same")),
      baselineContent: [value],
      currentContent: [value],
      serializedXml: emitted(""),
    }),
  ).toBeNull();
});

test("block deletion or movement cannot unbalance a formerly balanced comment range", () => {
  const first = block("11111111", "first");
  const second = block("22222222", "second");
  const start = `<q:p><q:commentRangeStart q:id="7"/></q:p>`;
  const end = `<q:p><q:commentRangeEnd q:id="7"/></q:p>`;
  const sourceXml = story(start + end);
  for (const currentContent of [[first], [second, first]]) {
    expect(
      buildStoryBlockReplay({
        sourceXml,
        baselineContent: [first, second],
        currentContent,
        serializedXml: emitted(currentContent.map(() => emittedParagraph("ignored")).join("")),
      }),
    ).toBeNull();
  }
});

const replaceFirstText = (value: BlockContent): void => {
  switch (value.type) {
    case "paragraph": {
      const run = value.content.find((entry) => entry.type === "run");
      const text = run?.content.find((entry) => entry.type === "text");
      if (!text) throw new Error("Missing parsed fixture text");
      text.text = "new";
      return;
    }
    case "blockSdt":
    case "blockCustomXml": {
      const child = value.content.at(0);
      if (!child) throw new Error("Missing parsed fixture child");
      replaceFirstText(child);
      return;
    }
    case "table": {
      const child = value.rows.at(0)?.cells.at(0)?.content.at(0);
      if (!child) throw new Error("Missing parsed fixture cell");
      replaceFirstText(child);
      return;
    }
    case "preservedBlock":
    case "bookmarkStart":
    case "bookmarkEnd":
      throw new Error("Fixture is not editable text");
    default: {
      const unreachable: never = value;
      return unreachable;
    }
  }
};

test.each(["blockSdt", "blockCustomXml", "table"] as const)(
  "edited nested content retains sibling bytes and original $0 wrapper syntax",
  (kind) => {
    const sourceSibling = `<q:p>\n  <q:r><q:rPr><q:webHidden/></q:rPr><q:t>same</q:t></q:r>\n</q:p>`;
    const sourceContent = sourceParagraph("old") + "\n<!-- nested gap -->\t" + sourceSibling;
    const generatedContent = emittedParagraph("new") + emittedParagraph("same");
    const sourceContainer = {
      table: `<q:tbl>\n<q:tblPr><q:tblLook q:val='04A0'/></q:tblPr><q:tblGrid/><q:tr>\n<q:tc>\n<q:tcPr><q:tcW q:w='2400' q:type='dxa'/></q:tcPr>\n${sourceContent}\n</q:tc></q:tr></q:tbl>`,
      blockSdt: `<q:sdt><q:sdtPr><q:richText/></q:sdtPr><q:sdtContent>\n${sourceContent}\n</q:sdtContent></q:sdt>`,
      blockCustomXml: `<q:customXml>\n${sourceContent}\n</q:customXml>`,
    }[kind];
    const generatedContainer = {
      table: `<w:tbl><w:tblPr><w:tblLook w:val='04A0'/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:tcPr><w:tcW w:w='2400' w:type='dxa'/></w:tcPr>${generatedContent}</w:tc></w:tr></w:tbl>`,
      blockSdt: `<w:sdt><w:sdtPr><w:richText/></w:sdtPr><w:sdtContent>${generatedContent}</w:sdtContent></w:sdt>`,
      blockCustomXml: `<w:customXml>${generatedContent}</w:customXml>`,
    }[kind];
    const baseline = parseHeader(story(sourceContainer)).content.at(0);
    if (!baseline) throw new Error("Missing parsed fixture container");
    const current = structuredClone(baseline);
    replaceFirstText(current);
    const output = buildStoryBlockReplay({
      sourceXml: story(sourceContainer),
      baselineContent: [baseline],
      currentContent: [current],
      serializedXml: emitted(generatedContainer),
    });
    expect(output).toContain(sourceSibling);
    expect(output).toContain("\n<!-- nested gap -->\t");
    expect(output).toContain(`<w:t>new</w:t>`);
    expect(output).not.toContain("<q:t>old</q:t>");
    if (kind === "table")
      expect(output).toContain(`<q:tc>\n<q:tcPr><q:tcW q:w='2400' q:type='dxa'/></q:tcPr>\n`);
    if (kind === "blockSdt")
      expect(output).toContain(`<q:sdt><q:sdtPr><q:richText/></q:sdtPr><q:sdtContent>\n`);
  },
);

test("opaque preserved children and one-block AlternateContent projections retain authored ranges", () => {
  const opaqueXml = `<vendor:payload xmlns:vendor='urn:vendor'>opaque</vendor:payload>`;
  const opaque = { type: "preservedBlock", xml: opaqueXml } as const satisfies BlockContent;
  const alternateXml = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback>${sourceParagraph("same")}</mc:Fallback></mc:AlternateContent>`;
  const projected = block(undefined, "same");
  const old = block(undefined, "old");
  const changed = block(undefined, "new");
  const output = buildStoryBlockReplay({
    sourceXml: story(opaqueXml + alternateXml + sourceParagraph("old")),
    baselineContent: [opaque, projected, old],
    currentContent: [opaque, projected, changed],
    serializedXml: emitted(opaqueXml + emittedParagraph("same") + emittedParagraph("new")),
  });
  expect(output).toContain(opaqueXml + alternateXml);
  expect(output).toContain(`<w:t>new</w:t>`);
});

test("changed extension markup inherits local Ignorable tokens without altering an untouched alias block", () => {
  const old = block(undefined, "old");
  const changed = block(undefined, "new");
  const untouched = block(undefined, "same");
  const sibling = `<q:p>\n<q:r><q:t>same</q:t></q:r></q:p>`;
  const mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const extension = "http://schemas.microsoft.com/office/word/2023/wordml/word16du";
  const output = buildStoryBlockReplay({
    sourceXml: story(sourceParagraph("old") + sibling),
    baselineContent: [old, untouched],
    currentContent: [changed, untouched],
    serializedXml: `<w:hdr xmlns:w="${WORD}" xmlns:compat="${mc}" xmlns:newext="${extension}" compat:Ignorable="newext"><w:p newext:dateUtc="2026-10-03"><w:r><w:t>new</w:t></w:r></w:p>${emittedParagraph("same")}</w:hdr>`,
  });
  expect(output).toContain(sibling);
  expect(output).toContain(`compat:Ignorable="newext"`);
  expect(output).toContain(`xmlns:newext="${extension}"`);
  expect(output?.startsWith(`<q:hdr xmlns:q="${WORD}">`)).toBe(true);
});

test("changed Strict table widths use source-profile slot conversion", () => {
  const sourceXml = `<q:hdr xmlns:q="${STRICT_WORD}"><q:tbl><q:tblPr><q:tblW q:w="25%" q:type="pct"/></q:tblPr><q:tblGrid/><q:tr><q:tc><q:tcPr/><q:p/></q:tc></q:tr></q:tbl></q:hdr>`;
  const baseline = parseHeader(sourceXml).content.at(0);
  if (!baseline || baseline.type !== "table") throw new Error("Missing fixture table");
  const current = {
    ...baseline,
    formatting: { ...baseline.formatting, width: { type: "pct", value: 2500 } },
  } as const satisfies BlockContent;
  const output = buildStoryBlockReplay({
    sourceXml,
    baselineContent: [baseline],
    currentContent: [current],
    serializedXml: emitted(
      `<w:tbl><w:tblPr><w:tblW w:w="2500" w:type="pct"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>`,
    ),
  });
  expect(output).toContain(`w:w="50%"`);
  expect(output).toContain(`xmlns:w="${STRICT_WORD}"`);
  const reopened = parseHeader(output ?? "").content.at(0);
  if (!reopened || reopened.type !== "table") throw new Error("Missing reopened table");
  expect(reopened.formatting?.width).toEqual({ type: "pct", value: 2500 });
});

test("multiple compatibility aliases merge inherited and local Ignorable tokens", () => {
  const mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const output = buildStoryBlockReplay({
    sourceXml: story(sourceParagraph("old")),
    baselineContent: [block(undefined, "old")],
    currentContent: [block(undefined, "new")],
    serializedXml: `<w:hdr xmlns:w="${WORD}" xmlns:compat="${mc}" xmlns:a="urn:a" compat:Ignorable="a"><w:p xmlns:x="${mc}" xmlns:y="${mc}" xmlns:b="urn:b" y:Ignorable="b"><w:r><w:t>new</w:t></w:r></w:p></w:hdr>`,
  });
  expect(output).toContain(`y:Ignorable="a b"`);
  expect(output).toContain(`xmlns:a="urn:a"`);
});
