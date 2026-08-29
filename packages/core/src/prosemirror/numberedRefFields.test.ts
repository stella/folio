import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import type {
  ComplexField,
  Document,
  Hyperlink,
  Paragraph,
  ParagraphContent,
  SimpleField,
} from "../types/document";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { fromProseDoc } from "./conversion/fromProseDoc";
import { toProseDoc } from "./conversion/toProseDoc";
import { parseNumberedRefInstruction, resolveNumberedRefFields } from "./numberedRefFields";

const textRun = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});

function numberedParagraph(options: {
  text: string;
  level: number;
  marker: string;
  bookmark?: string;
  numId?: number;
  abstractNumId?: number;
  startOverride?: number;
  decimalLevels?: boolean;
  pPrMark?: Paragraph["pPrMark"];
}): Paragraph {
  const bookmarkContent: ParagraphContent[] = options.bookmark
    ? [
        { type: "bookmarkStart", id: options.level + 1, name: options.bookmark },
        textRun(options.text),
        { type: "bookmarkEnd", id: options.level + 1 },
      ]
    : [textRun(options.text)];
  return {
    type: "paragraph",
    formatting: { numPr: { numId: options.numId ?? 5, ilvl: options.level } },
    content: bookmarkContent,
    listRendering: {
      marker: options.marker,
      level: options.level,
      numId: options.numId ?? 5,
      abstractNumId: options.abstractNumId ?? 7,
      isBullet: false,
      numFmt: options.decimalLevels || options.level !== 2 ? "decimal" : "lowerLetter",
      levelNumFmts: options.decimalLevels
        ? ["decimal", "decimal", "decimal"]
        : ["decimal", "decimal", "lowerLetter"],
      levelStarts: [1, 1, 1],
      ...(options.startOverride === undefined ? {} : { startOverride: options.startOverride }),
    },
    ...(options.pPrMark ? { pPrMark: options.pPrMark } : {}),
  };
}

function refField(
  instruction: string,
  displayText: string,
  kind: "simple" | "complex" = "complex",
): SimpleField | ComplexField {
  if (kind === "simple") {
    return {
      type: "simpleField",
      instruction,
      fieldType: "REF",
      content: [textRun(displayText)],
    };
  }
  return {
    type: "complexField",
    instruction,
    fieldType: "REF",
    fieldCode: [],
    fieldResult: [textRun(displayText)],
  };
}

function documentOf(content: Paragraph[]): Document {
  return { package: { document: { content } } };
}

function fieldFallbacks(pmDoc: PMNode): string[] {
  return toFlowBlocks(pmDoc).flatMap((block) =>
    block.kind === "paragraph"
      ? block.runs.flatMap((run) => (run.kind === "field" ? [run.fallback ?? ""] : []))
      : [],
  );
}

function savedFieldTexts(pmDoc: PMNode): string[] {
  const model = fromProseDoc(pmDoc);
  return model.package.document.content.flatMap((block) => {
    if (block.type !== "paragraph") {
      return [];
    }
    return block.content.flatMap((content) => {
      let runs: SimpleField["content"] | ComplexField["fieldResult"] = [];
      if (content.type === "simpleField") {
        runs = content.content;
      } else if (content.type === "complexField") {
        runs = content.fieldResult;
      }
      return runs.flatMap((run) =>
        run.type === "run"
          ? run.content.flatMap((item) => (item.type === "text" ? [item.text] : []))
          : [],
      );
    });
  });
}

function resolvedFieldValues(pmDoc: PMNode): (string | undefined)[] {
  const resolved = resolveNumberedRefFields(pmDoc);
  const values: (string | undefined)[] = [];
  pmDoc.descendants((node) => {
    if (node.type.name !== "field" && node.type.name !== "structuredField") {
      return true;
    }
    values.push(resolved.get(node));
    return false;
  });
  return values;
}

describe("numbered REF instructions", () => {
  test("accepts one numbered switch and inert hyperlink/format switches", () => {
    expect(parseNumberedRefInstruction(' ref "clause-a" \\w \\h \\* MERGEFORMAT ')).toEqual({
      bookmark: "clause-a",
      numberSwitch: "w",
    });
    expect(parseNumberedRefInstruction("REF clause-a \\n")?.numberSwitch).toBe("n");
    expect(parseNumberedRefInstruction("REF clause-a \\r")?.numberSwitch).toBe("r");
  });

  test("fails closed for ambiguous or unsupported instructions", () => {
    expect(parseNumberedRefInstruction("REF clause-a \\n \\w")).toBeNull();
    expect(parseNumberedRefInstruction("REF clause-a \\t \\w")).toBeNull();
    expect(parseNumberedRefInstruction(`REF ${"a".repeat(257)} \\w`)).toBeNull();
  });
});

describe("numbered REF projection", () => {
  test("uses the same tracked-change counter stream as the visible marker", () => {
    const inserted = {
      kind: "ins",
      info: { id: 1, author: "Reviewer", date: "2026-08-29T10:00:00Z" },
    } as const;
    const deleted = {
      kind: "del",
      info: { id: 2, author: "Reviewer", date: "2026-08-29T10:01:00Z" },
    } as const;
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({ text: "Inserted one", level: 0, marker: "%1.", pPrMark: inserted }),
        numberedParagraph({ text: "Inserted two", level: 0, marker: "%1.", pPrMark: inserted }),
        numberedParagraph({
          text: "Deleted target",
          level: 0,
          marker: "%1.",
          bookmark: "deleted-target",
          pPrMark: deleted,
        }),
        { type: "paragraph", content: [refField(" REF deleted-target \\n ", "1")] },
      ]),
    );

    const blocks = toFlowBlocks(pmDoc);
    const markers = blocks.flatMap((block) =>
      block.kind === "paragraph" && block.attrs?.listMarker ? [block.attrs.listMarker] : [],
    );

    expect(markers).toEqual(["1.", "2.", "1."]);
    expect(resolvedFieldValues(pmDoc)).toEqual(["1"]);
    expect(fieldFallbacks(pmDoc)).toEqual(["1"]);
  });

  test("keeps full-number contexts separate for final and original tracked streams", () => {
    const inserted = {
      kind: "ins",
      info: { id: 1, author: "Reviewer", date: "2026-08-29T10:00:00Z" },
    } as const;
    const deleted = {
      kind: "del",
      info: { id: 2, author: "Reviewer", date: "2026-08-29T10:01:00Z" },
    } as const;
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({ text: "Original parent", level: 0, marker: "%1.", startOverride: 5 }),
        numberedParagraph({ text: "Inserted parent", level: 0, marker: "%1.", pPrMark: inserted }),
        numberedParagraph({
          text: "Deleted child",
          level: 1,
          marker: "%2.",
          bookmark: "deleted-child",
          pPrMark: deleted,
        }),
        { type: "paragraph", content: [refField(" REF deleted-child \\w ", "5.1")] },
      ]),
    );

    const blocks = toFlowBlocks(pmDoc);
    const markers = blocks.flatMap((block) =>
      block.kind === "paragraph" && block.attrs?.listMarker ? [block.attrs.listMarker] : [],
    );

    expect(markers).toEqual(["5.", "6.", "1."]);
    expect(resolvedFieldValues(pmDoc)).toEqual(["5.1"]);
    expect(fieldFallbacks(pmDoc)).toEqual(["5.1"]);
  });

  test("distinguishes current-level and full-context values", () => {
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({ text: "Article", level: 0, marker: "%1." }),
        numberedParagraph({ text: "Section", level: 1, marker: "%1.%2" }),
        numberedParagraph({ text: "Clause", level: 2, marker: "(%3)", bookmark: "clause" }),
        {
          type: "paragraph",
          content: [
            refField(" REF clause \\n ", "(a)", "simple"),
            refField(" REF clause \\w ", "1.1(a)"),
          ],
        },
      ]),
    );

    expect(fieldFallbacks(pmDoc)).toEqual(["(a)", "1.1(a)"]);
    expect(savedFieldTexts(pmDoc)).toEqual(["(a)", "1.1(a)"]);
  });

  test("resolves relative numbering from the REF field paragraph", () => {
    const sharedAncestorSource = numberedParagraph({
      text: "Source 4.3.1",
      level: 2,
      marker: "%1.%2.%3.",
      decimalLevels: true,
    });
    sharedAncestorSource.content.push(refField(" REF target \\r ", "5.2"));
    const sameParentSource = numberedParagraph({
      text: "Source 4.5.1",
      level: 2,
      marker: "%1.%2.%3.",
      decimalLevels: true,
    });
    sameParentSource.content.push(refField(" REF target \\r ", "2"));

    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({
          text: "Root 4",
          level: 0,
          marker: "%1.",
          startOverride: 4,
          decimalLevels: true,
        }),
        numberedParagraph({
          text: "Branch 4.3",
          level: 1,
          marker: "%1.%2.",
          startOverride: 3,
          decimalLevels: true,
        }),
        sharedAncestorSource,
        numberedParagraph({
          text: "Branch 4.4",
          level: 1,
          marker: "%1.%2.",
          decimalLevels: true,
        }),
        numberedParagraph({
          text: "Branch 4.5",
          level: 1,
          marker: "%1.%2.",
          decimalLevels: true,
        }),
        sameParentSource,
        numberedParagraph({
          text: "Target 4.5.2",
          level: 2,
          marker: "%1.%2.%3.",
          bookmark: "target",
          decimalLevels: true,
        }),
      ]),
    );

    expect(resolvedFieldValues(pmDoc)).toEqual(["5.2", "2"]);
    expect(fieldFallbacks(pmDoc)).toEqual(["5.2", "2"]);
    expect(savedFieldTexts(pmDoc)).toEqual(["5.2", "2"]);
  });

  test("uses full context when no numbered prefix is shared", () => {
    const source = numberedParagraph({
      text: "Source 4.1",
      level: 1,
      marker: "%1.%2.",
      decimalLevels: true,
    });
    source.content.push(refField(" REF target \\r ", "5.1"));
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({
          text: "Root 4",
          level: 0,
          marker: "%1.",
          startOverride: 4,
          decimalLevels: true,
        }),
        source,
        numberedParagraph({
          text: "Root 5",
          level: 0,
          marker: "%1.",
          decimalLevels: true,
        }),
        numberedParagraph({
          text: "Target 5.1",
          level: 1,
          marker: "%1.%2.",
          bookmark: "target",
          decimalLevels: true,
        }),
      ]),
    );

    expect(resolvedFieldValues(pmDoc)).toEqual(["5.1"]);
  });

  test("uses provable full context without a source numbering context", () => {
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({
          text: "Target",
          level: 0,
          marker: "%1.",
          bookmark: "target",
        }),
        { type: "paragraph", content: [refField(" REF target \\r ", "1")] },
      ]),
    );

    expect(resolvedFieldValues(pmDoc)).toEqual(["1"]);
    expect(fieldFallbacks(pmDoc)).toEqual(["1"]);
    expect(savedFieldTexts(pmDoc)).toEqual(["1"]);
  });

  test("derives relative suffixes from counter components across variant templates", () => {
    const source = numberedParagraph({
      text: "Source",
      level: 2,
      marker: "Clause %1.%2.%3.",
      decimalLevels: true,
    });
    source.content.push(refField(" REF target \\r ", "5-2"));
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({
          text: "Root",
          level: 0,
          marker: "Article %1.",
          startOverride: 4,
          decimalLevels: true,
        }),
        numberedParagraph({
          text: "Source branch",
          level: 1,
          marker: "Part %1/%2.",
          startOverride: 3,
          decimalLevels: true,
        }),
        source,
        numberedParagraph({
          text: "Target branch",
          level: 1,
          marker: "Division %1:%2.",
          numId: 6,
          startOverride: 5,
          decimalLevels: true,
        }),
        numberedParagraph({
          text: "Target",
          level: 2,
          marker: "§ %1-%2-%3.",
          bookmark: "target",
          numId: 6,
          startOverride: 2,
          decimalLevels: true,
        }),
      ]),
    );

    expect(resolvedFieldValues(pmDoc)).toEqual(["5-2"]);
    expect(fieldFallbacks(pmDoc)).toEqual(["5-2"]);
  });

  test("keeps current-level spelling around a same-parent relative number", () => {
    const source = numberedParagraph({ text: "Source", level: 2, marker: "(%3)" });
    source.content.push(refField(" REF target \\r ", "(b)"));
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({ text: "Root", level: 0, marker: "%1." }),
        numberedParagraph({ text: "Parent", level: 1, marker: "%1.%2." }),
        source,
        numberedParagraph({ text: "Target", level: 2, marker: "(%3)", bookmark: "target" }),
      ]),
    );

    expect(resolvedFieldValues(pmDoc)).toEqual(["(b)"]);
  });

  test("uses the layout caller's preseeded continuation state", () => {
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({ text: "Target", level: 0, marker: "%1.", bookmark: "target" }),
        { type: "paragraph", content: [refField(" REF target \\w ", "8")] },
      ]),
    );
    const listCounters = new Map([[5, [7, ...Array.from({ length: 8 }, () => Number.NaN)]]]);
    const blocks = toFlowBlocks(pmDoc, {
      listCounters,
      listAbstractCounters: new Map([[7, [7, ...Array.from({ length: 8 }, () => Number.NaN)]]]),
      listSeenNumIds: new Set(["5:0"]),
    });

    expect(
      blocks.flatMap((block) =>
        block.kind === "paragraph" && block.attrs?.listMarker ? [block.attrs.listMarker] : [],
      ),
    ).toEqual(["8."]);
    expect(
      blocks.flatMap((block) =>
        block.kind === "paragraph"
          ? block.runs.flatMap((run) => (run.kind === "field" ? [run.fallback] : []))
          : [],
      ),
    ).toEqual(["8"]);
    expect(listCounters.get(5)?.at(0)).toBe(8);
  });

  test("keeps bookmark targets nested in hyperlinks through layout and round-trip", () => {
    const hyperlink = {
      type: "hyperlink",
      href: "https://example.test",
      children: [
        textRun("before"),
        { type: "bookmarkStart", id: 41, name: "linked-target", colFirst: 2, colLast: 3 },
        textRun("inside"),
        { type: "bookmarkEnd", id: 41 },
        textRun("after"),
      ],
    } as const satisfies Hyperlink;
    const target = numberedParagraph({ text: "unused", level: 0, marker: "%1." });
    target.content = [hyperlink];
    const pmDoc = toProseDoc(
      documentOf([
        target,
        { type: "paragraph", content: [refField(" REF linked-target \\w ", "1")] },
      ]),
    );

    expect(fieldFallbacks(pmDoc)).toEqual(["1"]);
    expect(
      Array.from(
        { length: pmDoc.child(0).childCount },
        (_, index) => pmDoc.child(0).child(index).type.name,
      ),
    ).toEqual(["text", "bookmarkBoundary", "text", "bookmarkBoundary", "text"]);
    expect(
      pmDoc
        .child(0)
        .child(1)
        .marks.some((mark) => mark.type.name === "hyperlink"),
    ).toBe(true);
    const cloned = pmDoc.type.schema.nodeFromJSON(pmDoc.toJSON());
    expect(fieldFallbacks(cloned)).toEqual(["1"]);
    const saved = fromProseDoc(cloned);
    const savedTarget = saved.package.document.content.at(0);
    expect(savedTarget?.type === "paragraph" ? savedTarget.content : null).toEqual([hyperlink]);
    expect(fieldFallbacks(toProseDoc(saved))).toEqual(["1"]);
  });

  test("resolves a bookmark target nested in a simple-field hyperlink", () => {
    const field = {
      type: "simpleField",
      instruction: " DOCVARIABLE target ",
      fieldType: "DOCVARIABLE",
      content: [
        {
          type: "hyperlink",
          anchor: "field-target",
          children: [
            { type: "bookmarkStart", id: 42, name: "field-target" },
            textRun("inside"),
            { type: "bookmarkEnd", id: 42 },
          ],
        },
      ],
    } as const satisfies SimpleField;
    const target = numberedParagraph({ text: "unused", level: 0, marker: "%1." });
    target.content = [field];
    const pmDoc = toProseDoc(
      documentOf([
        target,
        { type: "paragraph", content: [refField(" REF field-target \\w ", "1")] },
      ]),
    );

    expect(resolvedFieldValues(pmDoc)).toEqual([undefined, "1"]);
    const cloned = pmDoc.type.schema.nodeFromJSON(JSON.parse(JSON.stringify(pmDoc.toJSON())));
    const savedTarget = fromProseDoc(cloned).package.document.content.at(0);
    expect(savedTarget?.type === "paragraph" ? savedTarget.content : null).toEqual([field]);
  });

  test("keeps authored caches when a target or calibration is unavailable", () => {
    const pmDoc = toProseDoc(
      documentOf([
        numberedParagraph({ text: "Target", level: 0, marker: "%1.", bookmark: "target" }),
        {
          type: "paragraph",
          content: [
            refField(" REF target \\w ", "Section 1"),
            refField(" REF missing \\w ", "missing-cache"),
            refField(" REF target \\t \\w ", "unsupported-cache"),
          ],
        },
      ]),
    );

    expect(fieldFallbacks(pmDoc)).toEqual(["Section 1", "missing-cache", "unsupported-cache"]);
    expect(savedFieldTexts(pmDoc)).toEqual(["Section 1", "missing-cache", "unsupported-cache"]);
  });

  test("the first bookmark declaration wins even when it is not numbered", () => {
    const pmDoc = toProseDoc(
      documentOf([
        {
          type: "paragraph",
          content: [
            { type: "bookmarkStart", id: 2, name: "duplicate" },
            textRun("First"),
            { type: "bookmarkEnd", id: 2 },
          ],
        },
        numberedParagraph({
          text: "Second",
          level: 0,
          marker: "%1.",
          bookmark: "duplicate",
        }),
        { type: "paragraph", content: [refField(" REF duplicate \\w ", "kept")] },
      ]),
    );

    expect(fieldFallbacks(pmDoc)).toEqual(["kept"]);
    expect(savedFieldTexts(pmDoc)).toEqual(["kept"]);
  });

  test("a calibrated field follows renumbering and refreshes the saved cache", () => {
    const initial = toProseDoc(
      documentOf([
        numberedParagraph({ text: "Before", level: 0, marker: "%1." }),
        numberedParagraph({ text: "Target", level: 0, marker: "%1.", bookmark: "target" }),
        { type: "paragraph", content: [refField(" REF target \\w ", "2")] },
      ]),
    );
    expect(fieldFallbacks(initial)).toEqual(["2"]);

    const cloned = initial.type.schema.nodeFromJSON(JSON.parse(JSON.stringify(initial.toJSON())));
    expect(fieldFallbacks(cloned)).toEqual(fieldFallbacks(initial));
    const target = cloned.child(1);
    const reference = cloned.child(2);
    const renumbered = cloned.type.create(cloned.attrs, [target, reference]);

    expect(fieldFallbacks(renumbered)).toEqual(["1"]);
    expect(savedFieldTexts(renumbered)).toEqual(["1"]);
  });
});
