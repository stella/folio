import { describe, expect, test } from "bun:test";

import { parseSettings } from "../../../settingsParser";
import {
  bookmarkIdsOf,
  complexFieldXml,
  fieldBookmarkOf,
  freezeFields,
  mintRefBookmarkName,
  refFieldInstruction,
  setUpdateFieldsOnOpen,
  wrapParagraphContentInBookmark,
} from "../builders";
import { createStoryState, readParagraph, walkBlocks } from "../story";
import {
  elementChildren,
  partPrefixes,
  type RawPart,
  readRawPart,
  spliceXml,
  wChild,
} from "../xml";
import { fieldRuns, NS, p, r, style, XML } from "./fixtures";

const part = (body: string): RawPart => {
  const raw = readRawPart(
    "word/document.xml",
    `${XML}<w:document ${NS}><w:body>${body}</w:body></w:document>`,
  );
  if (raw === null) {
    throw new Error("fixture is not XML");
  }
  return raw;
};

const prefixesOf = (raw: RawPart) => {
  const prefixes = partPrefixes(raw.root);
  if (prefixes === null) {
    throw new Error("fixture binds no WordprocessingML prefix");
  }
  return prefixes;
};

type ReadStory = {
  texts: string[];
  fields: { kind: string; instruction: string; result: string }[];
};

const readStory = (raw: RawPart): ReadStory => {
  const state = createStoryState();
  const texts: string[] = [];
  const body = wChild(raw.root, "body");
  if (body === null) {
    throw new Error("fixture has no body");
  }
  walkBlocks(body, {
    paragraph: (element) => {
      texts.push(readParagraph(element, state).text);
    },
    tableStart: () => {},
    cellStart: () => {},
    tableEnd: () => {},
    unread: () => {},
  });
  return {
    texts,
    fields: state.fields.map(({ kind, instruction, result }) => ({
      kind,
      instruction: instruction.trim(),
      result,
    })),
  };
};

const frozen = (
  body: string,
  kinds: readonly string[],
): ReadStory & {
  xml: string;
  counts: { kind: string; count: number }[];
} => {
  const raw = part(body);
  const result = freezeFields(raw, new Set(kinds));
  const xml = spliceXml(raw.xml, result.splices);
  if (xml === null) {
    throw new Error("freeze splices were refused");
  }
  const reread = readRawPart(raw.path, xml);
  if (reread === null) {
    throw new Error("frozen part is not XML");
  }
  return { ...readStory(reread), xml, counts: result.frozen };
};

describe("mintRefBookmarkName", () => {
  test("is `_Ref` and nine digits, stable per seed and unique in the set", () => {
    const taken = new Set<string>();
    const first = mintRefBookmarkName("p4", taken);
    const again = mintRefBookmarkName("p4", new Set());
    const collided = mintRefBookmarkName("p4", taken);

    expect(first).toMatch(/^_Ref\d{9}$/u);
    expect(again).toBe(first);
    expect(collided).toMatch(/^_Ref\d{9}$/u);
    expect(collided).not.toBe(first);
    expect(taken).toEqual(new Set([first, collided]));
  });
});

describe("wrapParagraphContentInBookmark", () => {
  const raw = part("");
  const prefixes = prefixesOf(raw);
  const wrap = (paragraphXml: string): string | null =>
    wrapParagraphContentInBookmark({
      paragraphXml,
      prefixes,
      id: 7,
      name: "_Ref000000007",
    });

  test("puts the bookmark after the paragraph properties", () => {
    expect(wrap(p(r("Target"), style("HouseClause")))).toBe(
      '<w:p><w:pPr><w:pStyle w:val="HouseClause"/></w:pPr><w:bookmarkStart w:id="7" w:name="_Ref000000007"/><w:r><w:t xml:space="preserve">Target</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>',
    );
  });

  test("wraps a paragraph without properties and an empty one", () => {
    expect(wrap(p(r("x")))).toBe(
      '<w:p><w:bookmarkStart w:id="7" w:name="_Ref000000007"/><w:r><w:t xml:space="preserve">x</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>',
    );
    expect(wrap("<w:p/>")).toBe(
      '<w:p><w:bookmarkStart w:id="7" w:name="_Ref000000007"/><w:bookmarkEnd w:id="7"/></w:p>',
    );
  });

  test("refuses markup that is not one paragraph", () => {
    expect(wrap(`${p(r("a"))}${p(r("b"))}`)).toBeNull();
    expect(wrap("<w:tbl/>")).toBeNull();
  });

  test("reports the ids a part already uses", () => {
    const used = part(
      p(
        `<w:bookmarkStart w:id="3" w:name="a"/>${r("x")}<w:bookmarkEnd w:id="3"/><w:bookmarkStart w:id="12" w:name="b"/><w:bookmarkEnd w:id="12"/>`,
      ),
    );

    expect(bookmarkIdsOf(used.root).toSorted((a, b) => a - b)).toEqual([
      3, 3, 12, 12,
    ]);
  });
});

describe("complexFieldXml", () => {
  test("writes a REF field the reader takes its cached result from", () => {
    const instruction = refFieldInstruction("_Ref000000007", "\\n \\t \\h");
    expect(instruction).toBe(" REF _Ref000000007 \\n \\t \\h ");

    const field = complexFieldXml({
      prefix: "w",
      instruction: instruction ?? "",
      resultRunsXml: r("XI", "<w:b/>"),
      runPropertiesXml: "<w:rPr><w:b/></w:rPr>",
    });
    const story = readStory(part(p(`${r("see article ")}${field}${r(".")}`)));

    expect(story.texts).toEqual(["see article XI."]);
    expect(story.fields).toEqual([
      {
        kind: "REF",
        instruction: "REF _Ref000000007 \\n \\t \\h",
        result: "XI",
      },
    ]);
    expect(fieldBookmarkOf(instruction ?? "")).toBe("_Ref000000007");
    expect(field.match(/<w:rPr><w:b\/><\/w:rPr>/gu)).toHaveLength(5);
  });

  test("refuses switches that are not switches", () => {
    expect(refFieldInstruction("_Ref1", '\\h "quoted text"')).toBeNull();
    expect(refFieldInstruction("_Ref1", "MERGEFORMAT")).toBeNull();
    expect(refFieldInstruction("_Ref1", "")).toBe(" REF _Ref1 ");
  });
});

describe("setUpdateFieldsOnOpen", () => {
  const settings = `${XML}<w:settings ${NS}><w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/><w:compat/><w:decimalSymbol w:val=","/></w:settings>`;

  test("inserts the flag where the schema orders it and leaves the rest", () => {
    const set = setUpdateFieldsOnOpen(settings, true);

    expect(set).not.toBeNull();
    expect(parseSettings(set).updateFields).toBe(true);
    const root = elementChildren(readRawPart("s", set ?? "")?.root);
    expect(root.map(({ name }) => name?.replace(/^.*:/u, ""))).toEqual([
      "zoom",
      "defaultTabStop",
      "updateFields",
      "compat",
      "decimalSymbol",
    ]);
    expect(set?.replace(/<folioW:updateFields[^>]*\/>/u, "")).toBe(settings);
  });

  test("is idempotent and clears by removing the flag", () => {
    const set = setUpdateFieldsOnOpen(settings, true) ?? "";

    expect(setUpdateFieldsOnOpen(set, true)).toBe(set);
    expect(setUpdateFieldsOnOpen(set, false)).toBe(settings);
    expect(setUpdateFieldsOnOpen(settings, false)).toBe(settings);
  });

  test("refuses a part that is not a settings part", () => {
    expect(setUpdateFieldsOnOpen(`<w:document ${NS}/>`, true)).toBeNull();
  });
});

describe("freezeFields", () => {
  test("replaces a complex field by its cached result", () => {
    const result = frozen(
      p(`${r("To: ")}${fieldRuns('FILLIN "Recipient"', r("Recipient name"))}`),
      ["FILLIN"],
    );

    expect(result.texts).toEqual(["To: Recipient name"]);
    expect(result.fields).toEqual([]);
    expect(result.xml).not.toContain("fldChar");
    expect(result.xml).not.toContain("instrText");
    expect(result.counts).toEqual([{ kind: "FILLIN", count: 1 }]);
  });

  test("unwraps a simple field and keeps other kinds live", () => {
    const result = frozen(
      p(
        `<w:fldSimple w:instr=" DATE \\@ &quot;d.M.yyyy&quot; ">${r("1.1.2020")}</w:fldSimple>${r(" page ")}${fieldRuns("PAGE", r("3"))}`,
      ),
      ["DATE", "TIME"],
    );

    expect(result.texts).toEqual(["1.1.2020 page 3"]);
    expect(result.fields.map(({ kind }) => kind)).toEqual(["PAGE"]);
    expect(result.xml).not.toContain("fldSimple");
    expect(result.counts).toEqual([{ kind: "DATE", count: 1 }]);
  });

  test("keeps a field nested in a frozen result", () => {
    const body = `${p(`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r("First\t")}${fieldRuns("PAGEREF _Toc1 \\h", r("2"))}`)}${p(`${r("Second\t")}${fieldRuns("PAGEREF _Toc2 \\h", r("3"))}<w:r><w:fldChar w:fldCharType="end"/></w:r>`)}`;
    const result = frozen(body, ["TOC"]);

    expect(result.texts).toEqual(["First\t2", "Second\t3"]);
    expect(result.fields.map(({ kind, result: text }) => [kind, text])).toEqual(
      [
        ["PAGEREF", "2"],
        ["PAGEREF", "3"],
      ],
    );
    expect(result.counts).toEqual([{ kind: "TOC", count: 1 }]);
  });

  test("removes a field nested in a frozen field's code with the code", () => {
    const body = p(
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF </w:instrText></w:r>${fieldRuns("PAGE", r("1"))}<w:r><w:instrText> = 1 "first" "later" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r("first")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
    );

    const whole = frozen(body, ["IF"]);
    expect(whole.texts).toEqual(["first"]);
    expect(whole.fields).toEqual([]);
    expect(whole.xml).not.toContain("fldChar");

    // A field inside another field's code is that field's business.
    const inner = frozen(body, ["PAGE"]);
    expect(inner.fields.map(({ kind }) => kind)).toEqual(["IF", "PAGE"]);
    expect(inner.counts).toEqual([]);
  });

  test("freezes a field that never produced a result to nothing", () => {
    const result = frozen(
      p(
        `${r("a")}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> FILLIN "x" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>${r("b")}`,
      ),
      ["FILLIN"],
    );

    expect(result.texts).toEqual(["ab"]);
    expect(result.fields).toEqual([]);
  });
});
