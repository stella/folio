import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { createDocx } from "../../docx/rezip";
import type {
  Document,
  RunContent,
  SimpleField,
  TextFormatting,
  TrackedRunChange,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { schema } from "../schema";
import { TRACKED_RUN_INLINE_ATOM_DISPOSITIONS } from "./toProseDoc";

const REVISION_INFO = {
  id: 91,
  author: "Reviewer",
  date: "2026-09-09T00:00:00.000Z",
} as const;

const FORMATTED_RUN = {
  underline: { style: "single" },
} as const satisfies TextFormatting;

type AtomFixture = {
  name: string;
  content: RunContent;
  xmlTag: "br" | "tab";
};

const ATOM_FIXTURES = [
  {
    name: "hard break",
    content: { type: "break", breakType: "textWrapping" },
    xmlTag: "br",
  },
  {
    name: "tab",
    content: { type: "tab" },
    xmlTag: "tab",
  },
] as const satisfies readonly AtomFixture[];

const withMainContent = (content: Document["package"]["document"]["content"]): Document => {
  const document = createEmptyDocument();
  return {
    ...document,
    package: {
      ...document.package,
      document: {
        ...document.package.document,
        content,
      },
    },
  };
};

const reviewedAtomDocument = (atom: RunContent, type: "insertion" | "deletion"): Document => {
  const change: TrackedRunChange = {
    type,
    info: REVISION_INFO,
    content: [{ type: "run", formatting: FORMATTED_RUN, content: [atom] }],
  };
  return withMainContent([{ type: "paragraph", content: [change] }]);
};

const unreviewedAtomDocument = (atom: RunContent, formatting?: TextFormatting): Document =>
  withMainContent([
    {
      type: "paragraph",
      content: [
        {
          type: "run",
          ...(formatting ? { formatting } : {}),
          content: [atom],
        },
      ],
    },
  ]);

const reviewedFieldDocument = (field: SimpleField, type: "insertion" | "deletion"): Document =>
  withMainContent([
    {
      type: "paragraph",
      content: [
        {
          type,
          info: REVISION_INFO,
          content: [field],
        },
      ],
    },
  ]);

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const part = zip.file("word/document.xml");
  if (!part) {
    throw new Error("missing word/document.xml");
  }
  return part.async("text");
};

const elementCount = (xml: string, tag: AtomFixture["xmlTag"] | "fldSimple"): number =>
  xml.match(new RegExp(`<w:${tag}(?:[\\s/>])`, "gu"))?.length ?? 0;

const hasUnderline = (xml: string): boolean => /<w:u(?:[\s/>])/u.test(xml);

const resolveAll = async (
  buffer: ArrayBuffer,
  decision: "accept" | "reject",
  expectedType: "insertion" | "deletion",
): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  expect(reviewer.getChanges()).toEqual([
    expect.objectContaining({ id: REVISION_INFO.id, type: expectedType, text: "" }),
  ]);
  expect(decision === "accept" ? reviewer.acceptAll() : reviewer.rejectAll()).toBe(1);
  const resolved = await reviewer.toBuffer();
  const reopened = await FolioDocxReviewer.fromBuffer(resolved);
  expect(reopened.getChanges()).toHaveLength(0);
  return resolved;
};

describe("tracked run inline atom ownership", () => {
  test("classifies every inline atom exposed by the editor schema", () => {
    const schemaAtoms = Object.values(schema.nodes)
      .filter((node) => node.isInline && node.isAtom)
      .map(({ name }) => name)
      .toSorted();

    expect(Object.keys(TRACKED_RUN_INLINE_ATOM_DISPOSITIONS).toSorted()).toEqual(schemaAtoms);

    for (const [name, disposition] of Object.entries(TRACKED_RUN_INLINE_ATOM_DISPOSITIONS)) {
      const node = schema.nodes[name];
      if (!node) {
        throw new Error(`missing classified inline atom ${name}`);
      }
      if (disposition !== "carry") {
        continue;
      }
      expect(node.allowsMarkType(schema.marks["insertion"]!)).toBe(true);
      expect(node.allowsMarkType(schema.marks["deletion"]!)).toBe(true);
    }

    expect(TRACKED_RUN_INLINE_ATOM_DISPOSITIONS.field).toBe("field-carrier");
    expect(schema.nodes["field"]!.allowsMarkType(schema.marks["insertion"]!)).toBe(false);
    expect(schema.nodes["field"]!.allowsMarkType(schema.marks["deletion"]!)).toBe(false);
  });

  for (const { name, content, xmlTag } of ATOM_FIXTURES) {
    test(`accepts and rejects a formatted tracked ${name} after serialize and reopen`, async () => {
      for (const type of ["insertion", "deletion"] as const) {
        const pending = await createDocx(reviewedAtomDocument(content, type));
        const pendingXml = await documentXml(pending);
        expect(elementCount(pendingXml, xmlTag)).toBe(1);
        expect(pendingXml).toContain(`<w:${type === "insertion" ? "ins" : "del"} `);
        expect(hasUnderline(pendingXml)).toBe(true);

        const acceptedXml = await documentXml(await resolveAll(pending, "accept", type));
        const rejectedXml = await documentXml(await resolveAll(pending, "reject", type));
        expect(elementCount(acceptedXml, xmlTag)).toBe(type === "insertion" ? 1 : 0);
        expect(elementCount(rejectedXml, xmlTag)).toBe(type === "deletion" ? 1 : 0);
      }
    });

    test(`keeps unmarked and formatted ${name} runs through a no-edit save`, async () => {
      for (const formatting of [undefined, FORMATTED_RUN]) {
        const original = await createDocx(unreviewedAtomDocument(content, formatting));
        const reviewer = await FolioDocxReviewer.fromBuffer(original);
        expect(reviewer.getChanges()).toHaveLength(0);
        const savedXml = await documentXml(await reviewer.toBuffer());
        expect(elementCount(savedXml, xmlTag)).toBe(1);
        expect(hasUnderline(savedXml)).toBe(formatting !== undefined);
        expect(savedXml).not.toMatch(/<w:(?:ins|del)\b/u);
      }
    });
  }

  test("resolves the leaf-field carrier exception after serialize and reopen", async () => {
    const pending = await createDocx(
      reviewedFieldDocument(
        {
          type: "simpleField",
          instruction: " PAGE ",
          fieldType: "PAGE",
          content: [{ type: "run", content: [{ type: "text", text: "1" }] }],
        },
        "deletion",
      ),
    );

    expect(elementCount(await documentXml(pending), "fldSimple")).toBe(1);
    expect(
      elementCount(await documentXml(await resolveAll(pending, "accept", "deletion")), "fldSimple"),
    ).toBe(0);
    expect(
      elementCount(await documentXml(await resolveAll(pending, "reject", "deletion")), "fldSimple"),
    ).toBe(1);
  });

  test("resolves a structured-field carrier after serialize and reopen", async () => {
    const pending = await createDocx(
      reviewedFieldDocument(
        {
          type: "simpleField",
          instruction: " REF field-link \\h ",
          fieldType: "REF",
          content: [
            {
              type: "hyperlink",
              anchor: "field-link",
              children: [{ type: "run", content: [{ type: "text", text: "Field value" }] }],
            },
          ],
        },
        "insertion",
      ),
    );

    expect(elementCount(await documentXml(pending), "fldSimple")).toBe(1);
    expect(
      elementCount(
        await documentXml(await resolveAll(pending, "accept", "insertion")),
        "fldSimple",
      ),
    ).toBe(1);
    expect(
      elementCount(
        await documentXml(await resolveAll(pending, "reject", "insertion")),
        "fldSimple",
      ),
    ).toBe(0);
  });
});
