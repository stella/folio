import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { OP_STORIES, paragraphLogicalText, inlineLeafSpans } from "@stll/docx-core/ops";
import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
  ComplexField,
  NoteReferenceContent,
  Footnote,
  Endnote,
} from "../types/document";
import { toProseDoc, collectPairedBookmarkIds } from "../prosemirror/conversion/toProseDoc";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { projectCanonicalInline } from "./canonicalInlineProjection";
import { EditorState } from "prosemirror-state";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(30_000));

test("structured field payloads validate text without refusing authored instruction runs", () => {
  const field = {
    type: "complexField",
    instruction: "REF target",
    fieldType: "REF",
    fieldCode: [{ type: "run", content: [{ type: "instrText", text: "REF target" }] }],
    fieldResult: [run("cached")],
  } satisfies ComplexField;
  const source = {
    type: "paragraph",
    paraId: "12345678",
    content: [run("L"), field, run("R")],
  } satisfies Paragraph;
  expect(createCanonicalSession(documentFor(source)).isOk()).toBe(true);
  for (const text of ["\u0001", "\t", "\n"]) {
    field.fieldResult = [run(text)];
    const result = createCanonicalSession(documentFor(source));
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.message).toContain("valid XML characters");
  }
});

const run = (text: string) => ({ type: "run", content: [{ type: "text", text }] }) satisfies Run;
type RangeStart = Extract<ParagraphContent["type"], "bookmarkStart" | `${string}RangeStart`>;
const PAIRS = {
  bookmarkStart: [
    { type: "bookmarkStart", id: 1, name: "target" },
    { type: "bookmarkEnd", id: 1 },
  ],
  commentRangeStart: [
    { type: "commentRangeStart", id: 1 },
    { type: "commentRangeEnd", id: 1 },
  ],
  moveFromRangeStart: [
    { type: "moveFromRangeStart", id: 1, name: "source", author: "Reviewer" },
    { type: "moveFromRangeEnd", id: 1 },
  ],
  moveToRangeStart: [
    { type: "moveToRangeStart", id: 1, name: "destination", author: "Reviewer" },
    { type: "moveToRangeEnd", id: 1 },
  ],
} as const satisfies Record<RangeStart, readonly [ParagraphContent, ParagraphContent]>;

const documentFor = (paragraph: Paragraph) =>
  ({
    package: {
      document: {
        content: [paragraph],
        comments: paragraph.content.flatMap((item) =>
          item.type === "commentRangeStart"
            ? [
                {
                  id: item.id,
                  author: "Reviewer",
                  content: [
                    {
                      type: "paragraph",
                      paraId: (0x20000000 + item.id).toString(16),
                      content: [run("Comment")],
                    } satisfies Paragraph,
                  ],
                },
              ]
            : [],
        ),
      },
    },
  }) satisfies Document;

test("generated collapsed and paired markers retain exact source gap ordinals", () => {
  const check = (pairs: (typeof PAIRS)[RangeStart][]) => {
    const content: ParagraphContent[] = [run("L")];
    const expected = [{ position: 1, zeroWidthBefore: 0 }];
    let position = 1;
    let ordinal = 0;
    for (const [index, pair] of pairs.entries()) {
      content.push({ ...pair[0], id: index + 1 }, { ...pair[1], id: index + 1 });
      if (pair[0].type === "bookmarkStart") {
        expected.push({ position: ++position, zeroWidthBefore: ++ordinal });
        expected.push({ position: ++position, zeroWidthBefore: ++ordinal });
      } else {
        ordinal += 2;
        expected.push({ position: ++position, zeroWidthBefore: ordinal });
      }
    }
    content.push(run("R"));
    const source = { type: "paragraph", paraId: "12345678", content } satisfies Paragraph;
    const document = documentFor(source);
    const native = toProseDoc(document).child(0);
    const mapped = projectCanonicalInline({
      source,
      paragraph: native,
      pairedBookmarkIds: collectPairedBookmarkIds([source]),
    }).unwrap();
    expect(mapped.text).toBe(paragraphLogicalText(source));
    expect(mapped.text).toBe("LR");
    expect(mapped.boundaries).toEqual([
      [{ position: 0, zeroWidthBefore: 0 }],
      expected,
      [{ position: position + 1, zeroWidthBefore: 0 }],
    ]);
    const projection = createCanonicalSession(document).unwrap().projection;
    for (const gap of expected) {
      const address = projection.addressAt(1 + gap.position).unwrap();
      expect(address.offset).toBe(1);
      expect(address.zeroWidthBefore).toBe(gap.zeroWidthBefore);
      expect(projection.positionAt(address).unwrap()).toBe(1 + gap.position);
    }
    for (let missing = 0; missing <= ordinal; missing++) {
      if (expected.some((gap) => gap.zeroWidthBefore === missing)) continue;
      const refused = projection.positionAt({
        story: OP_STORIES.MAIN,
        blockId: source.paraId,
        offset: 1,
        zeroWidthBefore: missing,
      });
      expect(refused.isErr()).toBe(true);
      if (refused.isErr()) expect(refused.error.reason).toBe("refused");
    }
  };
  for (const pair of Object.values(PAIRS)) check([pair]);
  check(Object.values(PAIRS));
  assertProperty(
    fc.property(
      fc.array(fc.constantFrom(...Object.values(PAIRS)), { minLength: 1, maxLength: 6 }),
      check,
    ),
    { seed: -2053065844, numRuns: 40 },
  );
});

test("generated nonempty bookmark and move ranges preserve native gap geometry", () => {
  const pairs = Object.values(PAIRS).filter((pair) => pair[0].type !== "commentRangeStart");
  assertProperty(
    fc.property(fc.constantFrom(...pairs), fc.constantFrom("a", "契", "e\u0301"), (pair, text) => {
      const source = {
        type: "paragraph",
        paraId: "12345678",
        content: [run("L"), pair[0], run(text), pair[1], run("R")],
      } satisfies Paragraph;
      const document = documentFor(source);
      const native = toProseDoc(document).child(0);
      const mapped = projectCanonicalInline({
        source,
        paragraph: native,
        pairedBookmarkIds: collectPairedBookmarkIds([source]),
      }).unwrap();
      expect(mapped.text).toBe(paragraphLogicalText(source));
      expect(native.content.size).toBe(mapped.text.length + 2);
      expect(mapped.boundaries.at(1)).toEqual([
        { position: 1, zeroWidthBefore: 0 },
        { position: 2, zeroWidthBefore: 1 },
      ]);
      expect(mapped.boundaries.at(1 + text.length)).toEqual([
        { position: 2 + text.length, zeroWidthBefore: 0 },
        { position: 3 + text.length, zeroWidthBefore: 1 },
      ]);
      const projection = createCanonicalSession(document).unwrap().projection;
      for (const [offset, gaps] of mapped.boundaries.entries()) {
        for (const gap of gaps) {
          const address = projection.addressAt(1 + gap.position).unwrap();
          expect(address.offset).toBe(offset);
          expect(address.zeroWidthBefore ?? 0).toBe(gap.zeroWidthBefore);
          expect(projection.positionAt(address).unwrap()).toBe(1 + gap.position);
        }
      }
    }),
    { seed: 197, numRuns: 30 },
  );
});

test("nonempty comment ranges map their mark transitions to source gap ordinals", () => {
  const contents = [
    [run("L"), PAIRS.commentRangeStart[0], run("x"), PAIRS.commentRangeStart[1], run("R")],
  ] satisfies ParagraphContent[][];
  for (const content of contents) {
    const source = { type: "paragraph", paraId: "12345678", content } satisfies Paragraph;
    const document = documentFor(source);
    expect(createCanonicalSession(document).isOk()).toBe(true);
    const mapped = projectCanonicalInline({
      source,
      paragraph: toProseDoc(document).child(0),
      pairedBookmarkIds: new Set(),
    }).unwrap();
    expect(mapped.boundaries).toEqual([
      [{ position: 0, zeroWidthBefore: 0 }],
      [{ position: 1, zeroWidthBefore: 1 }],
      [{ position: 2, zeroWidthBefore: 1 }],
      [{ position: 3, zeroWidthBefore: 0 }],
    ]);
  }
});

const commentRangeContent = [
  run("L"),
  PAIRS.commentRangeStart[0],
  run("x"),
  PAIRS.commentRangeStart[1],
  run("R"),
] satisfies ParagraphContent[];

test("a comment range carried by native comment marks activates and keeps its boundaries through edits", () => {
  const source = {
    type: "paragraph",
    paraId: "12345678",
    content: commentRangeContent,
  } satisfies Paragraph;
  const session = createCanonicalSession(documentFor(source)).unwrap();
  // The range is a mark on "x": the native paragraph holds only "LxR" (positions 1-3).
  for (const edit of [
    { from: 3, to: 4, text: "Q", expected: "LxQ" },
    { from: 2, to: 3, text: "yz", expected: "LyzQ" },
  ]) {
    const state = EditorState.create({ doc: session.projection.doc });
    const commit = session.prepareReplace(state, { from: edit.from, to: edit.to, text: edit.text });
    publishCanonicalProjection({ session, state, commit: commit.unwrap() }).unwrap();
    const paragraph = session.document.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") panic("The edit lost its paragraph");
    expect(paragraphLogicalText(paragraph)).toBe(edit.expected);
    const boundaries = inlineLeafSpans(paragraph.content).flatMap(({ node, before }) =>
      node.type === "commentRangeStart" || node.type === "commentRangeEnd"
        ? [{ type: node.type, offset: before.offset }]
        : [],
    );
    // The comment keeps bracketing the text between "L" and the trailing character.
    expect(boundaries).toEqual([
      { type: "commentRangeStart", offset: 1 },
      { type: "commentRangeEnd", offset: edit.expected.length - 1 },
    ]);
  }
});

test("source markers with no native boundary receive typed activation refusals", () => {
  // Without a comment entry the range has no native mark, so its markers stay unmappable.
  const source = {
    type: "paragraph",
    paraId: "12345678",
    content: commentRangeContent,
  } satisfies Paragraph;
  const document = documentFor(source);
  document.package.document.comments = [];
  const result = createCanonicalSession(document);
  expect(result.isErr()).toBe(true);
  if (result.isErr()) expect(result.error.reason).toBe("refused");
});

test("generated erased containers keep unique source seams beside text and collapsed markers", () => {
  const empty = (kind: "hyperlink" | "bidi" | "smartTag" | "customXml"): ParagraphContent => {
    switch (kind) {
      case "hyperlink":
        return { type: "hyperlink", href: "https://empty.example/", children: [] };
      case "bidi":
        return { type: "inlineWrapper", kind, control: "override", direction: "rtl", content: [] };
      case "smartTag":
      case "customXml":
        return { type: "inlineWrapper", kind, element: "value", content: [] };
    }
  };
  const check = (kind: "hyperlink" | "bidi" | "smartTag" | "customXml", count: number) => {
    const containers = Array.from({ length: count }, () => empty(kind));
    const pair = PAIRS.bookmarkStart;
    const variants = [
      [...containers, run("LR")],
      [run("L"), ...containers, run("R")],
      [run("LR"), ...containers],
      [run("L"), pair[0], ...containers, pair[1], run("R")],
      [
        run("L"),
        { type: "inlineWrapper", kind: "smartTag", element: "outer", content: containers },
        run("R"),
      ],
      [run("L"), ...containers, ...PAIRS.moveFromRangeStart, run("R")],
      [run("L"), ...PAIRS.moveFromRangeStart, ...containers, run("R")],
      [run("L"), PAIRS.moveFromRangeStart[0], empty("bidi"), PAIRS.moveFromRangeStart[1], run("R")],
    ] satisfies ParagraphContent[][];
    for (const content of variants) {
      const source = { type: "paragraph", paraId: "12345678", content } satisfies Paragraph;
      const document = documentFor(source);
      const session = createCanonicalSession(document).unwrap();
      const native = session.projection.doc.child(0);
      if (content.some((item) => item.type === "moveFromRangeStart")) {
        const nativeKinds: string[] = [];
        native.forEach((node) => nativeKinds.push(node.type.name));
        expect(nativeKinds).toContain("rangeAnchor");
      }
      const mapped = projectCanonicalInline({
        source,
        paragraph: native,
        pairedBookmarkIds: collectPairedBookmarkIds([source]),
      }).unwrap();
      expect(mapped.text).toBe("LR");
      const zeroWidth = (paragraph: Paragraph) =>
        inlineLeafSpans(paragraph.content)
          .filter(({ before, after }) => before.offset === after.offset)
          .map(({ node }) => node);
      const preserved = session.document.package.document.content.at(0);
      expect(preserved?.type).toBe("paragraph");
      if (preserved?.type !== "paragraph") return;
      expect(zeroWidth(preserved)).toEqual(zeroWidth(source));
      for (const [offset, gaps] of mapped.boundaries.entries()) {
        expect(new Set(gaps.map(({ position }) => position)).size).toBe(gaps.length);
        for (const gap of gaps) {
          const address = session.projection.addressAt(1 + gap.position).unwrap();
          expect(address.offset).toBe(offset);
          expect(address.zeroWidthBefore ?? 0).toBe(gap.zeroWidthBefore);
          expect(session.projection.positionAt(address).unwrap()).toBe(1 + gap.position);
        }
        const maxOrdinal = Math.max(
          0,
          ...inlineLeafSpans(content)
            .filter(({ after }) => after.offset === offset)
            .map(({ after }) => after.zeroWidthBefore),
        );
        for (let ordinal = 0; ordinal <= maxOrdinal; ordinal += 1) {
          const position = session.projection.positionAt({
            story: OP_STORIES.MAIN,
            blockId: source.paraId,
            offset,
            zeroWidthBefore: ordinal,
          });
          expect(position.isOk()).toBe(
            gaps.some(({ zeroWidthBefore }) => zeroWidthBefore === ordinal),
          );
        }
      }
    }
  };
  for (const kind of ["hyperlink", "bidi", "smartTag", "customXml"] as const) check(kind, 2);
  assertProperty(
    fc.property(
      fc.constantFrom("hyperlink", "bidi", "smartTag", "customXml"),
      fc.integer({ min: 1, max: 5 }),
      check,
    ),
    { seed: -1348404097, numRuns: 30 },
  );
});

test("adjacent note-reference carriers preserve every distinct source occurrence", () => {
  const check = (kind: NoteReferenceContent["type"], count: number, ids: readonly number[]) => {
    const references = Array.from(
      { length: count },
      (_, index) =>
        ({
          type: kind,
          id: ids.at(index % ids.length) ?? panic("Adjacent reference fixture has no note id."),
        }) satisfies NoteReferenceContent,
    );
    const paragraph = {
      type: "paragraph",
      paraId: "12345678",
      content: [run("L"), { type: "run", content: references }, run("R")],
    } satisfies Paragraph;
    const noteParagraph = (id: number) =>
      ({
        type: "paragraph",
        paraId: (0x20000000 + id).toString(16),
        content: [run("Note")],
      }) satisfies Paragraph;
    const document = {
      package: {
        document: { content: [paragraph] },
        ...(kind === "footnoteRef"
          ? {
              footnotes: [...new Set(ids)].map(
                (id) => ({ type: "footnote", id, content: [noteParagraph(id)] }) satisfies Footnote,
              ),
            }
          : {
              endnotes: [...new Set(ids)].map(
                (id) => ({ type: "endnote", id, content: [noteParagraph(id)] }) satisfies Endnote,
              ),
            }),
      },
    } satisfies Document;
    const session = createCanonicalSession(document).unwrap();
    const native = session.projection.doc.child(0);
    const mapped = projectCanonicalInline({
      source: paragraph,
      paragraph: native,
      pairedBookmarkIds: collectPairedBookmarkIds([paragraph]),
    }).unwrap();
    expect(mapped.text).toBe(`L${"\uFFFC".repeat(count)}R`);
    const renderedNodes: string[] = [];
    native.forEach((node) => {
      if (node.marks.some(({ type }) => type.name === "footnoteRef"))
        renderedNodes.push(node.text ?? "");
    });
    expect(renderedNodes).toEqual(references.map((reference) => String(reference.id)));
    const occurrenceIds = new Set<string>();
    native.forEach((node) => {
      const mark = node.marks.find(({ type }) => type.name === "footnoteRef");
      if (mark) occurrenceIds.add(mark.attrs["occurrenceId"]);
    });
    expect(occurrenceIds.size).toBe(count);
    let physical = 2;
    for (const [index, reference] of references.entries()) {
      const offset = index + 1;
      expect(
        session.projection
          .positionAt({ story: OP_STORIES.MAIN, blockId: paragraph.paraId, offset })
          .unwrap(),
      ).toBe(physical);
      expect(session.projection.addressAt(physical).unwrap().offset).toBe(offset);
      for (let interior = 1; interior < String(reference.id).length; interior += 1)
        expect(session.projection.addressAt(physical + interior).isErr()).toBe(true);
      physical += String(reference.id).length;
    }
    expect(session.projection.addressAt(physical).unwrap().offset).toBe(count + 1);
    expect(
      session.projection
        .positionAt({ story: OP_STORIES.MAIN, blockId: paragraph.paraId, offset: count + 1 })
        .unwrap(),
    ).toBe(physical);
  };
  for (const kind of ["footnoteRef", "endnoteRef"] as const)
    for (const count of [2, 3])
      for (const ids of [[1], [11], [111], [1, 11, 111]]) check(kind, count, ids);
  assertProperty(
    fc.property(
      fc.constantFrom("footnoteRef", "endnoteRef"),
      fc.integer({ min: 2, max: 6 }),
      fc.constantFrom([1], [11], [111], [1, 11, 111]),
      check,
    ),
    { seed: 197, numRuns: 40 },
  );
});

test("adjacent field atoms retain separate source-unit boundaries", () => {
  const fields = [
    { type: "simpleField", instruction: "PAGE", fieldType: "PAGE", content: [run("123")] },
    {
      type: "complexField",
      instruction: "PAGE",
      fieldType: "PAGE",
      fieldCode: [{ type: "run", content: [{ type: "instrText", text: "PAGE" }] }],
      fieldResult: [run("123")],
    },
  ] satisfies ParagraphContent[];
  for (const field of fields)
    for (const count of [2, 3]) {
      const source = {
        type: "paragraph",
        paraId: "12345678",
        content: [run("L"), ...Array.from({ length: count }, () => field), run("R")],
      } satisfies Paragraph;
      const projection = createCanonicalSession(documentFor(source)).unwrap().projection;
      expect(projection.doc.child(0).childCount).toBe(count + 2);
      for (let offset = 0; offset <= count + 2; offset += 1) {
        const address = { story: OP_STORIES.MAIN, blockId: source.paraId, offset };
        expect(projection.addressAt(projection.positionAt(address).unwrap()).unwrap().offset).toBe(
          offset,
        );
      }
    }
});

test("story-owned eligibility preserves unpaired and cross-paragraph bookmark source gaps", () => {
  const cases = [
    { placement: "root", partner: "absent" },
    { placement: "wrapper", partner: "absent" },
    { placement: "root", partner: "otherParagraph" },
    { placement: "wrapper", partner: "otherParagraph" },
    { placement: "hyperlink", partner: "otherParagraph" },
    { placement: "revision", partner: "otherParagraph" },
  ] as const;
  const check = (endpoint: "start" | "end", { placement, partner }: (typeof cases)[number]) => {
    const pair = PAIRS.bookmarkStart;
    const marker = endpoint === "start" ? pair[0] : pair[1];
    const opposite = endpoint === "start" ? pair[1] : pair[0];
    const wrap = (item: ParagraphContent): ParagraphContent => {
      switch (placement) {
        case "root":
          return item;
        case "wrapper":
          return { type: "inlineWrapper", kind: "smartTag", element: "value", content: [item] };
        case "hyperlink":
          return { type: "hyperlink", href: "https://example.test/", children: [item] };
        case "revision":
          return {
            type: "insertion",
            info: { id: item.type === "bookmarkStart" ? 44 : 45, author: "Reviewer" },
            content: [item],
          };
      }
    };
    const source = {
      type: "paragraph",
      paraId: "12345678",
      content: [run("L"), wrap(marker), run("R")],
    } satisfies Paragraph;
    const other = {
      type: "paragraph",
      paraId: "22345678",
      content: [run("P"), wrap(opposite), run("Q")],
    } satisfies Paragraph;
    const pairedContent = endpoint === "start" ? [source, other] : [other, source];
    const content = partner === "absent" ? [source] : pairedContent;
    const document = { package: { document: { content } } } satisfies Document;
    const eligible = collectPairedBookmarkIds(content);
    expect(eligible.has(marker.id)).toBe(partner === "otherParagraph");
    const session = createCanonicalSession(document).unwrap();
    const sourceIndex = content.indexOf(source);
    const native = session.projection.doc.child(sourceIndex);
    const start = sourceIndex === 0 ? 1 : 1 + session.projection.doc.child(0).nodeSize;
    const mapped = projectCanonicalInline({
      source,
      paragraph: native,
      pairedBookmarkIds: eligible,
    }).unwrap();
    const represented =
      partner === "otherParagraph" || placement === "hyperlink" || placement === "revision";
    expect(native.content.size).toBe(represented ? 3 : 2);
    const current = session.projection.paragraph(source.paraId);
    expect(current).toBeDefined();
    if (!current) panic("Bookmark fixture lost its source paragraph.");
    expect(
      inlineLeafSpans(current.source.content)
        .filter(({ node }) => node.type === marker.type)
        .map(({ node }) => node),
    ).toEqual([marker]);
    for (const [offset, gaps] of mapped.boundaries.entries())
      for (const gap of gaps) {
        const address = session.projection.addressAt(start + gap.position).unwrap();
        expect(address.offset).toBe(offset);
        expect(address.zeroWidthBefore ?? 0).toBe(gap.zeroWidthBefore);
        expect(session.projection.positionAt(address).unwrap()).toBe(start + gap.position);
      }
    const before = session.projection.positionAt({
      story: OP_STORIES.MAIN,
      blockId: source.paraId,
      offset: 1,
      zeroWidthBefore: 0,
    });
    expect(before.isOk()).toBe(represented);
    const after = session.projection
      .positionAt({ story: OP_STORIES.MAIN, blockId: source.paraId, offset: 1, zeroWidthBefore: 1 })
      .unwrap();
    expect(session.projection.addressAt(after).unwrap().zeroWidthBefore).toBe(1);
  };
  for (const endpoint of ["start", "end"] as const)
    for (const scenario of cases) check(endpoint, scenario);
  assertProperty(fc.property(fc.constantFrom("start", "end"), fc.constantFrom(...cases), check), {
    seed: -1133252633,
    numRuns: 40,
  });
});
