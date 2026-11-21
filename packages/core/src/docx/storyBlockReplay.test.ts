import { expect, spyOn, test } from "bun:test";

import * as docx from "@stll/docx-core";
import * as xml from "@stll/docx-core/xml";
import * as xmlLimits from "@stll/docx-core/xml/limits";

import type { BlockContent } from "../types/document";
import { parseHeader } from "./headerFooterParser";
import {
  buildStoryBlockReplay,
  createReplayFragmentValidation,
  materializeReplayFragment,
} from "./storyBlockReplay";
import * as capture from "./verbatimCapture";
import { getChildElements, getNamespaceUri } from "./xmlParser";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
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
const sourceParagraph = (text: string) =>
  `<q:p><q:r><q:t>${text}</q:t></q:r></q:p>`;
const emittedParagraph = (text: string) =>
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

test("materializing a Strict fragment needs neither a full token list nor source ranges", () => {
  const fragment = '<w:p example="a > b"><w:r><w:t>changed</w:t></w:r></w:p>';
  const parsed = xml.parseStreamingXmlWithSourceRanges(emitted(fragment));
  if (parsed.status !== "parsed") throw new Error("Expected a generated story");
  const root = getChildElements(parsed.value).at(0);
  const element = getChildElements(root).at(0);
  if (!element) throw new Error("Expected a generated paragraph");

  const matchAll = String.prototype.matchAll;
  let fragmentTokenScans = 0;
  const tokens = spyOn(String.prototype, "matchAll").mockImplementation(
    function (this: string, ...args: Parameters<typeof matchAll>) {
      if (String(this) === fragment) fragmentTokenScans += 1;
      return matchAll.apply(this, args);
    },
  );
  const readsWithRanges = spyOn(xml, "parseStreamingXmlWithSourceRanges");
  try {
    const materialized = materializeReplayFragment({
      xml: fragment,
      element,
      sourceNamespace: STRICT_WORD,
    });
    expect(materialized).toContain(`xmlns:w="${STRICT_WORD}"`);
    expect(materialized).toContain("<w:t>changed</w:t>");
    expect(fragmentTokenScans).toBe(0);
    expect(readsWithRanges).toHaveBeenCalledTimes(0);
  } finally {
    readsWithRanges.mockRestore();
    tokens.mockRestore();
  }
});

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
      serializedXml: emitted(
        emittedParagraph("same") + emittedParagraph("same"),
      ),
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
    story(
      firstXml +
        " " +
        secondXml +
        `<w:p xmlns:w="${WORD}"><w:r><w:t>same</w:t></w:r></w:p>`,
    ),
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
      emittedParagraph("third") +
        emittedParagraph("first") +
        emittedParagraph("third"),
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
  expect(
    output?.startsWith(
      `<q:hdr xmlns:q='${STRICT_WORD}' xmlns:w="urn:foreign">`,
    ),
  ).toBe(true);
  const parsed = xml.parseStreamingXmlWithSourceRanges(output ?? "");
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
        serializedXml: emitted(
          currentContent.map(() => emittedParagraph("ignored")).join(""),
        ),
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
    const sourceContent =
      sourceParagraph("old") + "\n<!-- nested gap -->\t" + sourceSibling;
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
      expect(output).toContain(
        `<q:tc>\n<q:tcPr><q:tcW q:w='2400' q:type='dxa'/></q:tcPr>\n`,
      );
    if (kind === "blockSdt")
      expect(output).toContain(
        `<q:sdt><q:sdtPr><q:richText/></q:sdtPr><q:sdtContent>\n`,
      );
  },
);

test("opaque preserved children and one-block AlternateContent projections retain authored ranges", () => {
  const opaqueXml = `<vendor:payload xmlns:vendor='urn:vendor'>opaque</vendor:payload>`;
  const opaque = {
    type: "preservedBlock",
    xml: opaqueXml,
  } as const satisfies BlockContent;
  const alternateXml = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback>${sourceParagraph("same")}</mc:Fallback></mc:AlternateContent>`;
  const projected = block(undefined, "same");
  const old = block(undefined, "old");
  const changed = block(undefined, "new");
  const output = buildStoryBlockReplay({
    sourceXml: story(opaqueXml + alternateXml + sourceParagraph("old")),
    baselineContent: [opaque, projected, old],
    currentContent: [opaque, projected, changed],
    serializedXml: emitted(
      opaqueXml + emittedParagraph("same") + emittedParagraph("new"),
    ),
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
  const extension =
    "http://schemas.microsoft.com/office/word/2023/wordml/word16du";
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
  if (!baseline || baseline.type !== "table")
    throw new Error("Missing fixture table");
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
  if (!reopened || reopened.type !== "table")
    throw new Error("Missing reopened table");
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

const COMPAT = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const escapedNamespace = "urn:attributes?a=1&amp;b=2";

test("materialize replay escapes a shared sibling suffix once and validates every distinct fragment", () => {
  const fragments = Array.from(
    { length: 64 },
    (_, index) =>
      `<w:p data="${index} > 0"><w:r><w:t>${index}</w:t></w:r></w:p>`,
  );
  const parsed = xml.parseStreamingXmlWithSourceRanges(
    `<w:hdr xmlns:w="${WORD}" xmlns:mc="${COMPAT}" xmlns:x="${escapedNamespace}" mc:Ignorable="x x">${fragments.join("")}</w:hdr>`,
  );
  if (parsed.status !== "parsed")
    throw new Error("Expected generated siblings");
  const root = getChildElements(parsed.value).at(0);
  if (root === undefined) throw new Error("Expected generated root");
  const children = getChildElements(root);
  expect(children).toHaveLength(fragments.length);
  expect(new Set(children.map((child) => child.namespaceScope)).size).toBe(1);
  const validation = createReplayFragmentValidation();
  const escapes = spyOn(docx, "escapeXmlAttribute");
  const checks = spyOn(capture, "isSafeCapturedXmlDocument");
  try {
    for (const [index, fragment] of fragments.entries()) {
      const element = children[index];
      if (!element) throw new Error("Expected matching generated child");
      const result = materializeReplayFragment({
        xml: fragment,
        element,
        sourceNamespace: WORD,
        generatedRoot: root,
        validation,
      });
      const suffix = ` mc:Ignorable="x" xmlns:w="${WORD}" xmlns:mc="${COMPAT}" xmlns:x="${escapedNamespace}"`;
      expect(result).toBe(fragment.replace(' > 0">', ` > 0"${suffix}>`));
    }
    expect(escapes).toHaveBeenCalledTimes(4);
    expect(checks).toHaveBeenCalledTimes(fragments.length);
    const element = children.at(0);
    if (!element) throw new Error("Expected generated child");
    expect(
      materializeReplayFragment({
        xml: "<w:p><w:r></w:p>",
        element,
        sourceNamespace: WORD,
        generatedRoot: root,
        validation,
      }),
    ).toBeNull();
    expect(escapes).toHaveBeenCalledTimes(4);
    expect(checks).toHaveBeenCalledTimes(fragments.length + 1);
  } finally {
    checks.mockRestore();
    escapes.mockRestore();
  }
});

test.each([WORD, STRICT_WORD])(
  "replay suffix preserves local scopes, exact lexical compatibility, and root changes (%s)",
  (sourceNamespace) => {
    const fragments = [
      "<w:p/>",
      '<w:p xmlns:x="urn:shadow" mc:Ignorable="x y"/>',
      '<w:p mc:Ignorable="y"/>',
      "<w:p data='tail'/>",
    ];
    const parse = (ignorable: string) => {
      const parsed = xml.parseStreamingXmlWithSourceRanges(
        `<w:hdr xmlns:w="${WORD}" xmlns:alt="${COMPAT}" xmlns:mc="${COMPAT}" xmlns:x="${escapedNamespace}" xmlns:y="urn:y" alt:Ignorable="${ignorable}">${fragments.join("")}</w:hdr>`,
      );
      if (parsed.status !== "parsed")
        throw new Error("Expected generated story");
      const root = getChildElements(parsed.value).at(0);
      if (root === undefined) throw new Error("Expected generated root");
      return root;
    };
    const firstRoot = parse("x x");
    const secondRoot = parse("y");
    const validation = createReplayFragmentValidation();
    const children = getChildElements(firstRoot);
    const entries = fragments.map((raw, index) => ({
      raw,
      index,
      root: firstRoot,
    }));
    // Separate raw bytes still take the original lexical rewrite, even when
    // the element supplied by the caller has no own Ignorable attribute.
    entries.push({
      raw: "<w:p alt:Ignorable='old'/>",
      index: 0,
      root: firstRoot,
    });
    entries.push({ raw: fragments[0] ?? "", index: 0, root: secondRoot });
    entries.push({ raw: fragments[0] ?? "", index: 0, root: firstRoot });
    const escapes = spyOn(docx, "escapeXmlAttribute");
    try {
      const expectedEscapeCounts = [6, 5, 6, 0, 0, 6, 6];
      for (const [caseIndex, { raw, index, root }] of entries.entries()) {
        const element = children[index];
        if (!element) throw new Error("Expected generated child");
        const options = {
          xml: raw,
          element,
          sourceNamespace,
          generatedRoot: root,
        };
        const expected = materializeReplayFragment(options);
        const before = escapes.mock.calls.length;
        expect(materializeReplayFragment({ ...options, validation })).toBe(
          expected,
        );
        if (!raw.includes("old")) expect(expected).not.toBeNull();
        if (sourceNamespace === WORD)
          expect(escapes.mock.calls.length - before).toBe(
            expectedEscapeCounts[caseIndex],
          );
      }
    } finally {
      escapes.mockRestore();
    }
  },
);

test("compact replay reads shared namespace markup once and still checks raw siblings", () => {
  const fragments = Array.from(
    { length: 64 },
    (_, index) => `<w:p z="${index}"><w:r><w:t>${index}</w:t></w:r></w:p>`,
  );
  const parsed = xml.parseStreamingXmlWithSourceRanges(
    `<w:hdr xmlns:w="${WORD}" xmlns:mc="${COMPAT}" xmlns:x="urn:example" mc:Ignorable="x">${fragments.join("")}</w:hdr>`,
  );
  if (parsed.status !== "parsed")
    throw new Error("Expected generated siblings");
  const generatedRoot = getChildElements(parsed.value).at(0);
  if (!generatedRoot) throw new Error("Expected story root");
  const children = getChildElements(generatedRoot);
  const validation = createReplayFragmentValidation();
  const compact = spyOn(capture, "isSafeCompactReplayFragment");
  const full = spyOn(capture, "isSafeCapturedXmlDocument");
  const reader = spyOn(xml, "readXml");
  try {
    for (const [index, fragment] of fragments.entries()) {
      const element = children[index];
      if (!element) throw new Error("Expected generated child");
      expect(
        materializeReplayFragment({
          xml: fragment,
          element,
          sourceNamespace: WORD,
          generatedRoot,
          validation,
        }),
      ).toBe(
        fragment.replace(
          ` z="${index}">`,
          ` z="${index}" mc:Ignorable="x" xmlns:w="${WORD}" xmlns:mc="${COMPAT}" xmlns:x="urn:example">`,
        ),
      );
    }
    expect(compact).toHaveBeenCalledTimes(fragments.length);
    expect(full).toHaveBeenCalledTimes(0);
    expect(reader).toHaveBeenCalledTimes(fragments.length + 1);
    const reads = reader.mock.calls.map(([raw]) => raw);
    expect(reads.at(0)).toContain(' xmlns:x="urn:example"');
    expect(reads.slice(1)).toEqual(fragments);

    const element = children.at(0);
    if (!element) throw new Error("Expected generated child");
    for (const [raw, accepted] of [
      ["<w:p><w:r></w:p>", false],
      ["<w:p><no:unbound/></w:p>", false],
      ["<w:p><w:bad:name/></w:p>", false],
      ['<w:p xmlns:other="urn:other"><other:p/></w:p>', true],
      [`<w:p xmlns:w="${WORD}"/>`, false],
      ['<w:p mc:Ignorable="x"/>', false],
      ["<w:p>&amp;</w:p>", true],
      ["<w:p><?named harmless?></w:p>", true],
      ["<w:p>".repeat(100) + "</w:p>".repeat(100), false],
    ] as const) {
      const options = {
        xml: raw,
        element,
        sourceNamespace: WORD,
        generatedRoot,
      };
      const expected = materializeReplayFragment(options);
      expect(expected !== null).toBe(accepted);
      const before = full.mock.calls.length;
      expect(materializeReplayFragment({ ...options, validation })).toBe(
        expected,
      );
      expect(full.mock.calls.length - before).toBe(1);
    }
  } finally {
    reader.mockRestore();
    full.mockRestore();
    compact.mockRestore();
  }
});

test("compact proof charges suffix UTF-8 bytes and its validated attribute count", () => {
  const namespace = ` xmlns:w="${WORD}" xmlns:other="urn:é"`;
  const suffix = capture.prepareCompactReplaySuffix(namespace);
  if (suffix === undefined) throw new Error("Expected prepared replay suffix");
  const raw = '<w:p data="é"/>';
  const options = { xml: raw, suffix };
  const scanner = xmlLimits.assertXmlResourceLimits;
  const resources = spyOn(xmlLimits, "assertXmlResourceLimits");
  try {
    expect(capture.isSafeCompactReplayFragment(options)).toBe(true);
    expect(resources.mock.calls).toEqual([
      [
        {
          xml: raw,
          limits: {
            ...xmlLimits.FOLIO_XML_RESOURCE_LIMITS,
            maxBytes:
              xmlLimits.FOLIO_XML_RESOURCE_LIMITS.maxBytes -
              new TextEncoder().encode(namespace).length,
            maxAttributesPerPart:
              xmlLimits.FOLIO_XML_RESOURCE_LIMITS.maxAttributesPerPart - 2,
          },
        },
      ],
    ]);
    resources.mockImplementation((input) =>
      scanner({
        ...input,
        limits: {
          ...xmlLimits.FOLIO_XML_RESOURCE_LIMITS,
          maxAttributesPerPart: 0,
        },
      }),
    );
    expect(capture.isSafeCompactReplayFragment(options)).toBe(false);
  } finally {
    resources.mockRestore();
  }
});
