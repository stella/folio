import { describe, expect, test } from "bun:test";
import { Window } from "happy-dom";
import { DOMParser, DOMSerializer, type Mark } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";

import { parseParagraph } from "../../docx/paragraphParser";
import { serializeParagraph } from "../../docx/serializer/paragraphSerializer";
import { parseXmlDocument } from "../../docx/xmlParser";
import type { Paragraph } from "../../types/document";
import { acceptAIEditRevision, rejectAIEditRevision } from "../commands/comments";
import { extractTrackedChanges } from "../utils/extractTrackedChanges";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const REVISIONS = [
  { tag: "ins", textTag: "t" },
  { tag: "del", textTag: "delText" },
  { tag: "moveFrom", textTag: "delText" },
  { tag: "moveTo", textTag: "t" },
] as const;

type Revision = (typeof REVISIONS)[number];

const parseParagraphXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml);
  if (root === null) {
    throw new Error("Failed to parse the nested revision fixture");
  }
  return parseParagraph(root, null, null, null, null, null);
};

const documentWith = (paragraph: Paragraph) => ({
  package: { document: { content: [paragraph] } },
});

const throughEditor = (paragraph: Paragraph): Paragraph => {
  const source = documentWith(paragraph);
  const saved = fromProseDoc(toProseDoc(source), source).package.document.content.at(0);
  if (saved?.type !== "paragraph") {
    throw new Error("The editor round trip lost its paragraph");
  }
  return saved;
};

const revisionXml = (revision: Revision, id: number, author: string, content: string): string =>
  `<w:${revision.tag} w:id="${id}" w:author="${author}" w:date="2026-09-22T10:00:00Z">` +
  `${content}</w:${revision.tag}>`;

const runXml = (textTag: Revision["textTag"], text: string): string =>
  `<w:r><w:${textTag}>${text}</w:${textTag}></w:r>`;

const nestedParagraphXml = (
  outer: Revision,
  inner: Revision,
  outerIndex: number,
  innerIndex: number,
): string => {
  const outerId = 100 + outerIndex;
  const innerId = 200 + innerIndex;
  const innerXml = revisionXml(
    inner,
    innerId,
    `Inner ${inner.tag}`,
    runXml(inner.textTag, `inner-${inner.tag}`),
  );
  const outerContent =
    runXml(outer.textTag, `outer-${outer.tag}-before`) +
    innerXml +
    runXml(outer.textTag, `outer-${outer.tag}-after`);
  return (
    `<w:p xmlns:w="${W}">` +
    revisionXml(outer, outerId, `Outer ${outer.tag}`, outerContent) +
    "</w:p>"
  );
};

const saveThroughEditor = (xml: string): string =>
  serializeParagraph(throughEditor(parseParagraphXml(xml)));

const rebindWordprocessingNamespace = (xml: string): string =>
  xml.replace("<w:p", `<w:p xmlns:w="${W}"`);

const CASES = REVISIONS.flatMap((outer, outerIndex) =>
  REVISIONS.map((inner, innerIndex) => ({ outer, inner, outerIndex, innerIndex })),
);

describe("nested tracked revision editor round trip", () => {
  test("preserves three revision levels", () => {
    const authored =
      `<w:p xmlns:w="${W}">` +
      revisionXml(
        REVISIONS[0],
        401,
        "First",
        revisionXml(
          REVISIONS[1],
          402,
          "Second",
          revisionXml(REVISIONS[2], 403, "Third", runXml("delText", "deep")),
        ),
      ) +
      "</w:p>";
    const once = saveThroughEditor(authored);
    expect(once).toBe(serializeParagraph(parseParagraphXml(authored)));
    expect(saveThroughEditor(rebindWordprocessingNamespace(once))).toBe(once);
  });

  test("preserves a transparent wrapper between revision levels", () => {
    const authored =
      `<w:p xmlns:w="${W}">` +
      revisionXml(
        REVISIONS[0],
        411,
        "Outer",
        runXml("t", "before") +
          `<w:dir w:val="ltr">${revisionXml(REVISIONS[1], 412, "Inner", runXml("delText", "inside"))}</w:dir>` +
          runXml("t", "after"),
      ) +
      "</w:p>";
    const once = saveThroughEditor(authored);
    expect(once).toBe(serializeParagraph(parseParagraphXml(authored)));
    expect(saveThroughEditor(rebindWordprocessingNamespace(once))).toBe(once);
  });

  test.each(CASES)(
    "preserves distinct $outer.tag outside $inner.tag",
    ({ outer, inner, outerIndex, innerIndex }) => {
      const authored = nestedParagraphXml(outer, inner, outerIndex, innerIndex);
      const expected = serializeParagraph(parseParagraphXml(authored));
      const once = saveThroughEditor(authored);

      expect(once).toBe(expected);
    },
  );

  test.each(CASES)(
    "reaches a two-save fixed point for $outer.tag outside $inner.tag",
    ({ outer, inner, outerIndex, innerIndex }) => {
      const authored = nestedParagraphXml(outer, inner, outerIndex, innerIndex);
      const once = saveThroughEditor(authored);
      expect(saveThroughEditor(rebindWordprocessingNamespace(once))).toBe(once);
    },
  );
});

const SIMPLE_OUTER_ID = 301;
const SIMPLE_INNER_ID = 302;
const SIMPLE_NESTING =
  `<w:p xmlns:w="${W}">` +
  `<w:ins w:id="${SIMPLE_OUTER_ID}" w:author="Outer author">` +
  runXml("t", "outer-before") +
  `<w:del w:id="${SIMPLE_INNER_ID}" w:author="Inner author">` +
  runXml("delText", "inner") +
  "</w:del>" +
  runXml("t", "outer-after") +
  "</w:ins></w:p>";

const resolvedXml = (
  mode: "accept" | "reject",
  revisionId: number,
  xml = SIMPLE_NESTING,
): string => {
  const paragraph = parseParagraphXml(xml);
  const source = documentWith(paragraph);
  const view = {
    state: EditorState.create({ doc: toProseDoc(source) }),
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  const command = mode === "accept" ? acceptAIEditRevision : rejectAIEditRevision;
  expect(command(revisionId)(view.state, view.dispatch)).toBe(true);
  const saved = fromProseDoc(view.state.doc, source).package.document.content.at(0);
  if (saved?.type !== "paragraph") {
    throw new Error("Resolving a nested revision lost its paragraph");
  }
  return serializeParagraph(saved);
};

const textIn = (xml: string): string =>
  [...xml.matchAll(/<w:(?:t|delText)(?: [^>]*)?>([^<]*)<\/w:(?:t|delText)>/gu)]
    .map(([, text]) => text ?? "")
    .join("");

describe("selectively resolving nested tracked revisions", () => {
  test("keeps revision ancestry through the editor DOM", () => {
    const source = documentWith(parseParagraphXml(SIMPLE_NESTING));
    const prose = toProseDoc(source);
    let original: Mark | undefined;
    prose.descendants((node) => {
      if (node.isText && node.text === "inner") {
        original = node.marks.find(
          ({ type }) => type.name === "insertion" || type.name === "deletion",
        );
      }
    });
    if (!original) {
      throw new Error("The inner revision mark is missing");
    }
    const document = new Window().document as unknown as globalThis.Document;
    const host = document.createElement("div");
    host.append(
      DOMSerializer.fromSchema(schema).serializeFragment(
        schema.node("paragraph", null, [schema.text("inner", [original])]).content,
        { document },
      ),
    );
    const parsed = DOMParser.fromSchema(schema).parse(host);
    const reparsed = parsed.firstChild?.firstChild?.marks.find(
      ({ type }) => type.name === "deletion",
    );
    expect(reparsed?.attrs["_docxRevisionAncestors"]).toEqual(
      original?.attrs["_docxRevisionAncestors"],
    );
  });

  test("lists both revisions for independent review", () => {
    const source = documentWith(parseParagraphXml(SIMPLE_NESTING));
    const state = EditorState.create({ doc: toProseDoc(source) });
    const entries = extractTrackedChanges(state).entries.filter(
      ({ type }) => type === "insertion" || type === "deletion",
    );
    expect(entries.map(({ revisionId }) => revisionId).sort()).toEqual([
      SIMPLE_OUTER_ID,
      SIMPLE_INNER_ID,
    ]);
  });

  test("keeps same-author nested revisions as separate review cards", () => {
    const sameAuthor = SIMPLE_NESTING.replace('w:author="Inner author"', 'w:author="Outer author"');
    const source = documentWith(parseParagraphXml(sameAuthor));
    const state = EditorState.create({ doc: toProseDoc(source) });
    const entries = extractTrackedChanges(state).entries.filter(
      ({ type }) => type === "insertion" || type === "deletion" || type === "replacement",
    );
    expect(entries.map(({ type, revisionId }) => [type, revisionId])).toEqual([
      ["insertion", SIMPLE_OUTER_ID],
      ["deletion", SIMPLE_INNER_ID],
    ]);
  });

  test("accepts the inner deletion without resolving the outer insertion", () => {
    const saved = resolvedXml("accept", SIMPLE_INNER_ID);
    expect(saved).toContain(`<w:ins w:id="${SIMPLE_OUTER_ID}"`);
    expect(saved).not.toContain(`<w:del w:id="${SIMPLE_INNER_ID}"`);
    expect(saved).not.toContain("inner");
    expect(textIn(saved)).toBe("outer-beforeouter-after");
  });

  test("rejects the inner deletion without resolving the outer insertion", () => {
    const saved = resolvedXml("reject", SIMPLE_INNER_ID);
    expect(saved).toContain(`<w:ins w:id="${SIMPLE_OUTER_ID}"`);
    expect(saved).not.toContain(`<w:del w:id="${SIMPLE_INNER_ID}"`);
    expect(textIn(saved)).toBe("outer-beforeinnerouter-after");
  });

  test("accepts the outer insertion without resolving the inner deletion", () => {
    const saved = resolvedXml("accept", SIMPLE_OUTER_ID);
    expect(saved).not.toContain(`<w:ins w:id="${SIMPLE_OUTER_ID}"`);
    expect(saved).toContain(`<w:del w:id="${SIMPLE_INNER_ID}"`);
    expect(saved).toContain("outer-before");
    expect(saved).toContain("inner");
    expect(saved).toContain("outer-after");
  });

  test("rejects the outer insertion including its nested deletion", () => {
    const saved = resolvedXml("reject", SIMPLE_OUTER_ID);
    expect(saved).not.toContain(`<w:ins w:id="${SIMPLE_OUTER_ID}"`);
    expect(saved).not.toContain(`<w:del w:id="${SIMPLE_INNER_ID}"`);
    expect(saved).not.toContain("outer-before");
    expect(saved).not.toContain("inner");
    expect(saved).not.toContain("outer-after");
  });

  test("rejects an outer deletion while keeping its inner insertion pending", () => {
    const xml =
      `<w:p xmlns:w="${W}">` +
      revisionXml(
        REVISIONS[1],
        SIMPLE_OUTER_ID,
        "Outer",
        revisionXml(REVISIONS[0], SIMPLE_INNER_ID, "Inner", runXml("t", "inner")),
      ) +
      "</w:p>";
    const saved = resolvedXml("reject", SIMPLE_OUTER_ID, xml);
    expect(saved).not.toContain(`<w:del w:id="${SIMPLE_OUTER_ID}"`);
    expect(saved).toContain(`<w:ins w:id="${SIMPLE_INNER_ID}"`);
    expect(textIn(saved)).toBe("inner");
  });
});
