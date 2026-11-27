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

/** What a reader sees: text by paragraph, an element as nesting. */
const shapeOf = (blocks: readonly BlockContent[]): unknown[] =>
  blocks.flatMap((block): unknown[] => {
    if (block.type === "blockCustomXml")
      return [{ element: shapeOf(block.content) }];
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
const decision = fc.constantFrom(
  REVISION_DECISIONS.ACCEPT,
  REVISION_DECISIONS.REJECT,
);

test("a suggested custom XML element resolves alike in every order, and after a reopen", async () => {
  const base = await opened();
  const cases: {
    inner: string[];
    offset: number;
    element: RevisionDecision;
    content: RevisionDecision;
  }[] = [];
  assertProperty(
    fc.property(
      fc.array(word, { minLength: 1, maxLength: 3 }),
      fc.integer({ min: 0, max: DESTINATION.length }),
      decision,
      decision,
      (inner, offset, element, content) => {
        cases.push({ inner, offset, element, content });
      },
    ),
    { id: "suggested custom XML element resolution orders", numRuns: 20 },
  );
  for (const { inner, offset, element, content } of cases) {
    const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset };
    const compiled = compileEditorIntent(base, {
      intent: {
        type: "pasteBlocks",
        from: at,
        to: at,
        blocks: [
          paragraph("Lead"),
          {
            type: "blockCustomXml",
            openingXml: '<w:customXml w:element="clause">',
            closingXml: "</w:customXml>",
            content: inner.map(paragraph),
          },
          paragraph("Trail"),
        ],
        openStart: 1,
        openEnd: 1,
      },
      mode: {
        type: "suggesting",
        revision: { id: 900, author: "R", date: "2026-01-02T00:00:00Z" },
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
    expect(revisionIdsOf(reopened).length).toBe(
      revisionIdsOf(suggested).length,
    );
    expect([...outcomes(reopened, { element, content })]).toEqual([...here]);
  }
});
