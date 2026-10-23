import { describe, expect, test } from "bun:test";

import type {
  ComplexField,
  Document,
  ParagraphContent,
  Run,
} from "../../types/document";
import { canonicalFieldNode } from "../fieldRepresentation";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const WORD_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const resultRuns = [
  {
    type: "run",
    content: [{ type: "text", text: "A" }],
    preservedAttributes: [
      { namespace: WORD_NAMESPACE, name: "rsidR", value: "00112233" },
    ],
  },
  {
    type: "run",
    content: [{ type: "text", text: "B" }],
    formatting: {},
    preservedAttributes: [
      { namespace: WORD_NAMESPACE, name: "rsidRPr", value: "00445566" },
    ],
  },
  {
    type: "run",
    content: [{ type: "text", text: "C" }],
    preservedAttributes: [
      { namespace: WORD_NAMESPACE, name: "rsidDel", value: "00ABCDEF" },
    ],
  },
] satisfies Run[];

const fieldCode = [
  {
    type: "run",
    content: [{ type: "instrText", text: " REF " }],
    preservedAttributes: [
      { namespace: WORD_NAMESPACE, name: "rsidR", value: "00778899" },
    ],
  },
  {
    type: "run",
    content: [{ type: "instrText", text: "target" }],
    formatting: {},
    preservedAttributes: [
      { namespace: WORD_NAMESPACE, name: "rsidRPr", value: "00AABBCC" },
    ],
  },
] satisfies ComplexField["fieldCode"];

const differentlyFormattedResultRuns = [
  {
    type: "run",
    content: [{ type: "text", text: "bold" }],
    formatting: { bold: true },
  },
  {
    type: "run",
    content: [{ type: "text", text: "italic" }],
    formatting: { italic: true },
  },
] satisfies Run[];

const resultRunsWithMetadata = (count: number): Run[] =>
  Array.from({ length: count }, (_, index): Run => ({
    type: "run",
    content: [{ type: "text", text: String.fromCharCode(65 + index) }],
    ...(index % 2 === 1 ? { formatting: {} } : {}),
    preservedAttributes: [
      {
        namespace: WORD_NAMESPACE,
        name: ["rsidR", "rsidRPr", "rsidDel"][index % 3] ?? "rsidR",
        value: (index + 1).toString(16).toUpperCase().padStart(8, "0"),
      },
    ],
  }));

type FieldContent = Extract<
  ParagraphContent,
  { type: "simpleField" | "complexField" }
>;

const documentWithField = (field: FieldContent): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", content: [field] }],
    },
  },
});

const firstField = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected a field paragraph");
  }
  const field = paragraph.content.at(0);
  if (field?.type !== "simpleField" && field?.type !== "complexField") {
    throw new Error("Expected a field");
  }
  return field;
};

const roundTrip = (field: FieldContent): ParagraphContent => {
  const source = documentWithField(field);
  return firstField(
    fromProseDoc(toProseDoc(source), source, {
      stylesheetSource: { type: "package" },
    }),
  );
};

const projectField = (field: FieldContent) => {
  const node = toProseDoc(documentWithField(field)).firstChild?.firstChild;
  if (!node) throw new Error("Expected a projected field node");
  return node;
};

describe("field result run metadata round-trip", () => {
  test("preserves metadata at every run boundary across field representations", () => {
    for (let count = 2; count <= 5; count += 1) {
      const runs = resultRunsWithMetadata(count);
      const simpleField = {
        type: "simpleField",
        instruction: " REF target ",
        fieldType: "REF",
        content: runs,
      } satisfies Extract<ParagraphContent, { type: "simpleField" }>;
      const complexField = {
        type: "complexField",
        instruction: " REF target ",
        fieldType: "REF",
        fieldCode,
        fieldResult: runs,
      } satisfies ComplexField;

      const restoredSimple = roundTrip(simpleField);
      const restoredComplex = roundTrip(complexField);
      expect(restoredSimple).toMatchObject({
        type: "simpleField",
        content: runs,
      });
      expect(restoredComplex).toMatchObject({
        type: "complexField",
        fieldResult: runs,
      });
    }
  });

  test("keeps metadata-bearing result runs structured through canonicalization", () => {
    for (const field of [
      {
        type: "simpleField",
        instruction: " REF target ",
        fieldType: "REF",
        content: resultRuns,
      },
      {
        type: "complexField",
        instruction: " REF target ",
        fieldType: "REF",
        fieldCode,
        fieldResult: resultRuns,
      },
    ] satisfies FieldContent[]) {
      const projected = projectField(field);
      expect(projected.type).toBe(schema.nodes.structuredField);
      expect(canonicalFieldNode(projected)).toBe(projected);
    }
  });

  test("keeps a plain single-run field atomic", () => {
    const source = {
      type: "simpleField",
      instruction: " PAGE ",
      fieldType: "PAGE",
      content: [{ type: "run", content: [{ type: "text", text: "1" }] }],
    } satisfies Extract<ParagraphContent, { type: "simpleField" }>;

    expect(projectField(source).type).toBe(schema.nodes.field);
  });

  test("retains native note carriers inside both field representations", () => {
    for (const content of [
      { type: "footnoteRef", id: 3 },
      { type: "endnoteRef", id: 4 },
      { type: "noteMarker", kind: "footnote" },
      { type: "noteMarker", kind: "endnote" },
    ] satisfies Run["content"]) {
      const run = { type: "run", content: [content] } satisfies Run;
      const simple = {
        type: "simpleField",
        instruction: " REF target ",
        fieldType: "REF",
        content: [run],
      } satisfies FieldContent;
      const complex = {
        type: "complexField",
        instruction: " REF target ",
        fieldType: "REF",
        fieldCode,
        fieldResult: [run],
      } satisfies FieldContent;

      expect(projectField(simple).type).toBe(schema.nodes.structuredField);
      expect(projectField(complex).type).toBe(schema.nodes.structuredField);
      expect(roundTrip(simple)).toMatchObject({ content: [run] });
      expect(roundTrip(complex)).toMatchObject({ fieldResult: [run] });
    }
  });

  test("retains distinct formatting across text-only result runs", () => {
    const simpleSource = {
      type: "simpleField",
      instruction: " REF target ",
      fieldType: "REF",
      content: differentlyFormattedResultRuns,
    } satisfies Extract<ParagraphContent, { type: "simpleField" }>;
    const complexSource = {
      type: "complexField",
      instruction: " REF target ",
      fieldType: "REF",
      fieldCode,
      fieldResult: differentlyFormattedResultRuns,
    } satisfies ComplexField;

    const projectedSimple = projectField(simpleSource);
    expect(projectedSimple.type).toBe(schema.nodes.structuredField);
    expect(canonicalFieldNode(projectedSimple)).toBe(projectedSimple);
    expect(roundTrip(simpleSource)).toMatchObject({
      type: "simpleField",
      content: differentlyFormattedResultRuns,
    });

    const projectedComplex = projectField(complexSource);
    expect(projectedComplex.type).toBe(schema.nodes.structuredField);
    expect(canonicalFieldNode(projectedComplex)).toBe(projectedComplex);
    expect(roundTrip(complexSource)).toMatchObject({
      type: "complexField",
      fieldCode,
      fieldResult: differentlyFormattedResultRuns,
    });
  });

  test("retains every simple-field result run's authored metadata", () => {
    const source = {
      type: "simpleField",
      instruction: " REF target ",
      fieldType: "REF",
      content: resultRuns,
    } satisfies Extract<ParagraphContent, { type: "simpleField" }>;

    const restored = roundTrip(source);

    expect(restored).toMatchObject({
      type: "simpleField",
      content: resultRuns,
    });
  });

  test("retains all complex-field result and code run metadata", () => {
    const source = {
      type: "complexField",
      instruction: " REF target ",
      fieldType: "REF",
      fieldCode,
      fieldResult: resultRuns,
    } satisfies ComplexField;

    const restored = roundTrip(source);

    expect(restored).toMatchObject({
      type: "complexField",
      fieldCode,
      fieldResult: resultRuns,
    });
  });

  test.each(["field", "field-and-text"] as const)(
    "keeps a complex field inside its comment range: %s",
    (coverage) => {
      const followingContent: Run[] =
        coverage === "field"
          ? []
          : [
              {
                type: "run",
                content: [{ type: "text", text: "Following text" }],
              },
            ];
      const expectedContent = [
        { type: "commentRangeStart", id: 17 },
        {
          type: "complexField",
          instruction: " REF target ",
          fieldType: "REF",
          fieldCode,
          fieldResult: resultRuns,
        },
        ...followingContent,
        { type: "commentRangeEnd", id: 17 },
        { type: "run", content: [{ type: "text", text: "Outside range" }] },
      ] satisfies ParagraphContent[];
      const source: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: expectedContent,
              },
            ],
          },
        },
      };

      const paragraph = fromProseDoc(toProseDoc(source), source, {
        stylesheetSource: { type: "package" },
      }).package.document.content.at(0);

      expect(paragraph?.type).toBe("paragraph");
      if (paragraph?.type !== "paragraph") {
        throw new Error("Expected a paragraph containing the field");
      }
      expect(paragraph.content).toMatchObject(expectedContent);
    },
  );
});
