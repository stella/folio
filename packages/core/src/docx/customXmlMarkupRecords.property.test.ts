/**
 * A suggested paste of a custom XML element, resolved: the element's record
 * of its own tags and the records of what it wraps leave one document
 * whichever is accepted or rejected first, and a save and reopen keeps every
 * record, so the reopened suggestion resolves to the same documents.
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOps,
  compileEditorIntent,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type {
  BlockContent,
  BlockCustomXml,
  Document,
  Paragraph,
  Table,
} from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(120_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

const paragraph = (text: string): Paragraph => ({
  type: "paragraph",
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const makeDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const DESTINATION = "before after";

const opened = async (): Promise<Document> =>
  parseDocx(
    await makeDocx(
      `<w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">${DESTINATION}</w:t></w:r></w:p>` +
        `<w:p w14:paraId="00000002"><w:r><w:t>next</w:t></w:r></w:p>`,
    ),
    { preloadFonts: false },
  );

const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]);

const table = (text: string): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [paragraph(text)] }],
    },
  ],
});

const tagged = (content: BlockContent[]): BlockCustomXml => ({
  type: "blockCustomXml",
  openingXml: '<w:customXml w:element="clause">',
  closingXml: "</w:customXml>",
  content,
});

/** One block an element holds: a paragraph, a table or another element. */
type Held =
  | { kind: "paragraph"; text: string }
  | { kind: "table"; text: string }
  | { kind: "element"; held: Held[] };

const blockOf = (held: Held): BlockContent => {
  if (held.kind === "paragraph") return paragraph(held.text);
  if (held.kind === "table") return table(held.text);
  return tagged(held.held.map(blockOf));
};

/** What a reader sees: text by paragraph, an element as nesting. */
const shapeOf = (blocks: readonly BlockContent[]): unknown[] =>
  blocks.flatMap((block): unknown[] => {
    if (block.type === "blockCustomXml")
      return [{ element: shapeOf(block.content) }];
    if (block.type === "table")
      return [
        {
          table: block.rows.map((row) =>
            row.cells.map((cell) => shapeOf(cell.content)),
          ),
        },
      ];
    if (block.type !== "paragraph") return [];
    const texts: string[] = [];
    const visit = (nodes: readonly unknown[]): void => {
      for (const node of nodes) {
        if (typeof node !== "object" || node === null) continue;
        if ("text" in node && typeof node.text === "string")
          texts.push(node.text);
        if ("content" in node && Array.isArray(node.content))
          visit(node.content);
      }
    };
    visit(block.content);
    return [texts.join("")];
  });

/** Every tracked record's revision id in the body, once. */
const revisionIdsOf = (document: Document): number[] => {
  const ids = new Set<number>();
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (
      "info" in value &&
      typeof value.info === "object" &&
      value.info !== null &&
      "id" in value.info &&
      typeof value.info.id === "number" &&
      "author" in value.info
    )
      ids.add(value.info.id);
    for (const field of Object.values(value)) visit(field);
  };
  visit(document.package.document.content);
  return [...ids];
};

/** Every tracked record's location, kind and metadata. */
const revisionFactsOf = (document: Document): string[] => {
  const facts: string[] = [];
  const visit = (value: unknown, path: string): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      value.forEach((item, index) => visit(item, `${path}[${index}]`));
      return;
    }
    if (
      "info" in value &&
      typeof value.info === "object" &&
      value.info !== null &&
      "id" in value.info &&
      "author" in value.info
    )
      facts.push(JSON.stringify({ path, info: value.info }));
    for (const [key, field] of Object.entries(value))
      visit(field, `${path}.${key}`);
  };
  visit(document.package.document.content, "content");
  return facts.toSorted();
};

const elementOf = (document: Document): BlockCustomXml => {
  const found = document.package.document.content.find(
    (block): block is BlockCustomXml => block.type === "blockCustomXml",
  );
  if (found === undefined) throw new TypeError("No custom XML element.");
  return found;
};

/** Each order of the two groups, and both at once when they agree. */
const outcomes = (
  document: Document,
  decisions: { element: RevisionDecision; content: RevisionDecision },
): Set<string> => {
  const own = elementOf(document).insertion?.info.id;
  if (own === undefined) throw new TypeError("The element records nothing.");
  const rest = revisionIdsOf(document).filter((id) => id !== own);
  const steps: [readonly number[], RevisionDecision][] = [
    [[own], decisions.element],
    [rest, decisions.content],
  ];
  const seen = new Set<string>();
  for (const order of [steps, steps.toReversed()]) {
    let current = document;
    for (const [ids, decision] of order)
      current = resolve(current, ids, decision);
    seen.add(JSON.stringify(shapeOf(current.package.document.content)));
  }
  if (decisions.element === decisions.content)
    seen.add(
      JSON.stringify(
        shapeOf(
          resolve(document, [own, ...rest], decisions.element).package.document
            .content,
        ),
      ),
    );
  return seen;
};

const word = fc.stringMatching(/^[a-z]{1,6}$/u);
const leaf: fc.Arbitrary<Held> = fc.oneof(
  word.map((text): Held => ({ kind: "paragraph", text })),
  word.map((text): Held => ({ kind: "table", text })),
);
/** An element's blocks: paragraphs, tables and nested elements of them. */
const heldArbitrary: fc.Arbitrary<Held[]> = fc.array(
  fc.oneof(
    { weight: 3, arbitrary: leaf },
    {
      weight: 1,
      arbitrary: fc
        .array(leaf, { minLength: 1, maxLength: 2 })
        .map((held): Held => ({ kind: "element", held })),
    },
  ),
  { minLength: 1, maxLength: 3 },
);
const decision = fc.constantFrom(
  REVISION_DECISIONS.ACCEPT,
  REVISION_DECISIONS.REJECT,
);

test("a suggested custom XML element resolves alike in every order, and after a reopen", async () => {
  const base = await opened();
  await assertProperty(
    fc.asyncProperty(
      heldArbitrary,
      fc.integer({ min: 0, max: DESTINATION.length }),
      decision,
      decision,
      async (inner, offset, element, content) => {
        const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset };
        const compiled = compileEditorIntent(base, {
          intent: {
            type: "pasteBlocks",
            from: at,
            to: at,
            blocks: [
              paragraph("Lead"),
              tagged(inner.map(blockOf)),
              paragraph("Trail"),
            ],
            openStart: 1,
            openEnd: 1,
          },
          mode: {
            type: "suggesting",
            revision: {
              id: 900,
              author: "R",
              date: "2026-01-02T00:00:00Z",
            },
            newIds: {
              revision: Array.from({ length: 60 }, (_, index) => 901 + index),
            },
          },
        });
        if (compiled.isErr()) throw compiled.error;
        const suggested = apply(base, compiled.value.ops);
        const here = outcomes(suggested, { element, content });
        expect(here.size).toBe(1);
        const reopened = await parseDocx(
          await repackDocx(suggested, { updateModifiedDate: false }),
          { preloadFonts: false },
        );
        expect(elementOf(reopened).insertion?.info).toEqual(
          elementOf(suggested).insertion?.info,
        );
        expect(revisionFactsOf(reopened)).toEqual(revisionFactsOf(suggested));
        expect([...outcomes(reopened, { element, content })]).toEqual([
          ...here,
        ]);
      },
    ),
    {
      id: "suggested custom XML element resolution orders",
      numRuns: 40,
    },
  );
});
