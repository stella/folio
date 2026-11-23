import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../types/document";
import { derivedFieldOps } from "./derivedFields";

const INSTRUCTION = "EVIDENCE invoice";

const documentWith = (paragraph: Paragraph): Document => ({
  package: { document: { content: [paragraph] } },
});

const linkedField = (text: string): Paragraph => ({
  type: "paragraph",
  paraId: "12345678",
  content: [
    {
      type: "simpleField",
      instruction: INSTRUCTION,
      fieldType: "UNKNOWN",
      content: [
        {
          type: "hyperlink",
          href: "https://example.com",
          children: [{ type: "run", content: [{ type: "text", text }] }],
        },
      ],
    },
  ],
});

const derive = (text: string, derived: string) =>
  derivedFieldOps(documentWith(linkedField(text)), {
    owns: (instruction) => instruction === INSTRUCTION,
    displayTexts: (instructions) => instructions.map(() => derived),
  });

test("a field's text reads through the hyperlink that holds it", () => {
  expect(derive("Exhibit 1", "Exhibit 1")).toEqual([]);
  expect(derive("Exhibit 2", "Exhibit 1")).toHaveLength(1);
  expect(derive("Exhibit 1", "")).toHaveLength(1);
});
