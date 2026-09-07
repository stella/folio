/**
 * The shapes a compared package has to have, read off its own XML.
 *
 * The round-trip self-check reads content: which words a view resolves to,
 * which cell they sit in. It cannot see a package that carries the right words
 * in markup the format does not allow — a row before the table's grid, a
 * hyperlink inside a revision, two revisions under one id — because parsing it
 * back gives the same content either way. These assertions look at the markup
 * instead.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { buildBodySequenceDocx, type BodyItem } from "./__fixtures__/body-sequence";
import { compareDocx } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2024-03-01T00:00:00.000Z" } as const;

const comparedDocumentXml = async (
  base: readonly BodyItem[],
  target: readonly BodyItem[],
): Promise<string> => {
  const result = await compareDocx(
    await buildBodySequenceDocx(base),
    await buildBodySequenceDocx(target),
    OPTIONS,
  );
  if (result.isErr()) {
    throw result.error;
  }
  const zip = await JSZip.loadAsync(result.value.buffer);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
};

const REVISION_ELEMENT =
  /<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange|cellIns|cellDel|cellMerge)\b[^>]*\bw:id="(\d+)"/gu;

const revisionIdsOf = (xml: string): string[] =>
  [...xml.matchAll(REVISION_ELEMENT)].map(([, , id]) => id ?? "");

const TAG = /<(\/?)([A-Za-z0-9_:.-]+)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/gu;

/** The element names directly under each `w:tbl`, outermost tables first. */
const tableChildNames = (xml: string): string[][] => {
  const stack: { name: string; children: string[] }[] = [];
  const tables: string[][] = [];
  for (const match of xml.matchAll(TAG)) {
    const [, closing, name = "", , selfClosing] = match;
    if (closing) {
      if (stack.at(-1)?.name === name) {
        stack.pop();
      }
      continue;
    }
    if (stack.at(-1)?.name === "w:tbl") {
      stack.at(-1)?.children.push(name);
    }
    if (selfClosing) {
      continue;
    }
    const frame = { name, children: [] as string[] };
    if (name === "w:tbl") {
      tables.push(frame.children);
    }
    stack.push(frame);
  }
  return tables;
};

/** Ancestor element names of every `w:hyperlink`, innermost last. */
const hyperlinkAncestors = (xml: string): string[][] => {
  const stack: string[] = [];
  const ancestors: string[][] = [];
  for (const match of xml.matchAll(TAG)) {
    const [, closing, name = "", , selfClosing] = match;
    if (closing) {
      if (stack.at(-1) === name) {
        stack.pop();
      }
      continue;
    }
    if (name === "w:hyperlink") {
      ancestors.push([...stack]);
    }
    if (!selfClosing) {
      stack.push(name);
    }
  }
  return ancestors;
};

const LINK = { text: "the guide", href: "https://example.invalid/guide" } as const;

const TABLE: BodyItem = {
  kind: "table",
  rows: [
    ["a1", "a2"],
    ["b1", "b2"],
    ["c1", "c2"],
  ],
};

describe("revision ids", () => {
  test("a word-level redline gives every emitted wrapper its own id", async () => {
    // One `replaceInBlock` allocates one id for its deletions and one for its
    // insertions, and the redline cuts each into several wrappers around the
    // words that survived. Every wrapper still needs an id of its own.
    const xml = await comparedDocumentXml(
      [{ kind: "paragraph", text: "alpha beta gamma delta epsilon zeta" }],
      [{ kind: "paragraph", text: "alpha BETA gamma delta EPSILON zeta" }],
    );
    const ids = revisionIdsOf(xml);
    expect(ids.length).toBeGreaterThan(2);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe("table placement", () => {
  test.each([
    {
      name: "an added table",
      base: [{ kind: "paragraph", text: "before" }] satisfies BodyItem[],
      target: [{ kind: "paragraph", text: "before" }, TABLE] satisfies BodyItem[],
    },
    {
      name: "a removed table",
      base: [{ kind: "paragraph", text: "before" }, TABLE] satisfies BodyItem[],
      target: [{ kind: "paragraph", text: "before" }] satisfies BodyItem[],
    },
    {
      name: "a row added at the top",
      base: [TABLE] satisfies BodyItem[],
      target: [{ kind: "table", rows: [["new1", "new2"], ...TABLE.rows] }] satisfies BodyItem[],
    },
    {
      name: "a row added in the middle",
      base: [TABLE] satisfies BodyItem[],
      target: [
        {
          kind: "table",
          rows: [TABLE.rows[0] ?? [], ["new1", "new2"], ...TABLE.rows.slice(1)],
        },
      ] satisfies BodyItem[],
    },
    {
      name: "a row added at the end",
      base: [TABLE] satisfies BodyItem[],
      target: [{ kind: "table", rows: [...TABLE.rows, ["new1", "new2"]] }] satisfies BodyItem[],
    },
    {
      name: "a row removed from the middle",
      base: [TABLE] satisfies BodyItem[],
      target: [
        { kind: "table", rows: [TABLE.rows[0] ?? [], TABLE.rows[2] ?? []] },
      ] satisfies BodyItem[],
    },
    {
      name: "the last row removed",
      base: [TABLE] satisfies BodyItem[],
      target: [{ kind: "table", rows: TABLE.rows.slice(0, 2) }] satisfies BodyItem[],
    },
  ])("$name keeps every row after the table's properties and grid", async ({ base, target }) => {
    const xml = await comparedDocumentXml(base, target);
    const tables = tableChildNames(xml);
    expect(tables.length).toBeGreaterThan(0);
    for (const children of tables) {
      expect(children.slice(0, 2)).toEqual(["w:tblPr", "w:tblGrid"]);
      expect(children.slice(2).every((name) => name === "w:tr")).toBe(true);
    }
    // A tracked row is marked inside its own `w:trPr`, never by wrapping the
    // row in a revision element.
    expect(/<w:(?:ins|del)\b[^>]*>\s*<w:tr\b/u.test(xml)).toBe(false);
  });
});

describe("hyperlink nesting", () => {
  const LINKED_PARAGRAPH = {
    kind: "paragraph",
    text: ["see ", LINK, " for details"],
  } as const satisfies BodyItem;

  test.each([
    {
      name: "a deleted hyperlink",
      base: [{ kind: "paragraph", text: "keep" }, LINKED_PARAGRAPH] satisfies BodyItem[],
      target: [{ kind: "paragraph", text: "keep" }] satisfies BodyItem[],
    },
    {
      name: "an inserted hyperlink",
      base: [{ kind: "paragraph", text: "keep" }] satisfies BodyItem[],
      target: [{ kind: "paragraph", text: "keep" }, LINKED_PARAGRAPH] satisfies BodyItem[],
    },
    {
      name: "text edited around a hyperlink",
      base: [LINKED_PARAGRAPH] satisfies BodyItem[],
      target: [{ kind: "paragraph", text: ["read ", LINK, " for more"] }] satisfies BodyItem[],
    },
  ])("$name keeps the revision inside the link", async ({ base, target }) => {
    const xml = await comparedDocumentXml(base, target);
    for (const ancestors of hyperlinkAncestors(xml)) {
      expect(ancestors).not.toContain("w:ins");
      expect(ancestors).not.toContain("w:del");
      expect(ancestors).not.toContain("w:moveFrom");
      expect(ancestors).not.toContain("w:moveTo");
    }
  });

  test("a deleted hyperlink's runs carry deleted text", async () => {
    const xml = await comparedDocumentXml(
      [{ kind: "paragraph", text: "keep" }, LINKED_PARAGRAPH],
      [{ kind: "paragraph", text: "keep" }],
    );
    expect(xml).toContain(`<w:hyperlink r:id="rId2"><w:del`);
    expect(/<w:hyperlink[^>]*><w:del\b[^>]*>(?:(?!<\/w:del>).)*<w:delText/su.test(xml)).toBe(true);
  });
});

describe("determinism", () => {
  test("two runs over the same inputs produce byte-identical packages", async () => {
    // `dcterms:modified` is the last clock the package holds: a save stamps it
    // from the wall clock, which makes two otherwise identical runs differ in
    // that part alone.
    const base = await buildBodySequenceDocx([
      { kind: "paragraph", text: "alpha beta gamma" },
      TABLE,
      { kind: "paragraph", text: ["see ", LINK, " for details"] },
    ]);
    const target = await buildBodySequenceDocx([
      { kind: "paragraph", text: "alpha BETA gamma" },
      TABLE,
      { kind: "paragraph", text: ["read ", LINK, " for more"] },
    ]);
    const [first, second] = await Promise.all([
      compareDocx(base, target, OPTIONS),
      compareDocx(base, target, OPTIONS),
    ]);
    if (first.isErr()) {
      throw first.error;
    }
    if (second.isErr()) {
      throw second.error;
    }
    expect(new Uint8Array(second.value.buffer)).toEqual(new Uint8Array(first.value.buffer));

    const zip = await JSZip.loadAsync(first.value.buffer);
    expect(await zip.file("docProps/core.xml")?.async("text")).toContain(
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${OPTIONS.timestamp}</dcterms:modified>`,
    );
  });
});
