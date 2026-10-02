import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph, Run } from "../../model/document";
import { storyParagraphs } from "../blocks";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  createEditorIntentIdAllocator,
} from "../editorIntent";
import { IDENTITY_SPACES, idKey, packageIdentityKeys, packageParagraphIds } from "../ids";
import { paragraphLength } from "../offsets";
import { filterNewIds } from "../plan";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";
import { documentArbitrary } from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

const paragraph = (index: number): Paragraph => ({
  type: "paragraph",
  paraId: (index + 1).toString(16).padStart(8, "0").toUpperCase(),
  content: [{ type: "run", content: [{ type: "text", text: "abc" }] }],
});

const replaceMainParagraph = (document: Document, replacement: Paragraph): Document => ({
  ...document,
  package: {
    ...document.package,
    document: {
      ...document.package.document,
      content: document.package.document.content.map((block) =>
        block.type === "paragraph" && block.paraId === replacement.paraId ? replacement : block,
      ),
    },
  },
});

const replaceHeaderIdentities = (
  document: Document,
  candidate: { newBlockId: string; revisionId: number },
): Document => ({
  ...document,
  package: {
    ...document.package,
    headers: new Map([
      [
        "rIdHeader",
        {
          type: "header",
          hdrFtrType: "default",
          content: [
            {
              type: "paragraph",
              paraId: candidate.newBlockId,
              content: [
                {
                  type: "insertion",
                  info: { id: candidate.revisionId, author: "Header" },
                  content: [{ type: "run", content: [{ type: "text", text: "header" }] }],
                },
              ],
            },
          ],
        },
      ],
    ]),
  },
});

describe("editor intent identity allocation", () => {
  test("filtering pools preserves supplied spaces and ordered identities", () => {
    assertProperty(
      fc.property(
        fc.option(fc.array(fc.nat(100)), { nil: undefined }),
        fc.option(fc.array(fc.nat(100)), { nil: undefined }),
        (revision, control) => {
          const newIds = {
            ...(revision === undefined ? {} : { revision }),
            ...(control === undefined ? {} : { control }),
          };
          const filtered = filterNewIds(
            {
              type: DOCUMENT_OP_TYPES.INSERT_TEXT,
              at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
              text: "x",
              newIds,
            },
            (_, id) => id % 2 === 0,
          );
          if (!("newIds" in filtered)) panic("Filtering lost the operation's identity pools.");
          expect(Object.keys(filtered.newIds ?? {})).toEqual(Object.keys(newIds));
          expect(filtered.newIds?.revision).toEqual(revision?.filter((id) => id % 2 === 0));
          expect(filtered.newIds?.control).toEqual(control?.filter((id) => id % 2 === 0));
        },
      ),
      { numRuns: 100 },
    );
  });

  test("single-character input requests bounded ids and reuses the package census", () => {
    let unrelatedContentReads = 0;
    const paragraphs = Array.from({ length: 2_000 }, (_, index) => {
      const value = paragraph(index);
      if (index === 0) return value;
      const content = value.content;
      Object.defineProperty(value, "content", {
        enumerable: true,
        get: () => {
          unrelatedContentReads += 1;
          return content;
        },
      });
      return value;
    });
    const document: Document = { package: { document: { content: paragraphs } } };
    const allocate = createEditorIntentIdAllocator();
    const intent = {
      type: "replaceText",
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      text: "x",
    } as const;
    const first = allocate(document, intent);
    const firstReads = unrelatedContentReads;
    expect(firstReads).toBeGreaterThanOrEqual(1_999);
    expect(first.newIds.revision?.length ?? 0).toBeLessThanOrEqual(16);
    expect(first.newIds.control?.length ?? 0).toBeLessThanOrEqual(16);
    const second = allocate(document, intent);
    expect(unrelatedContentReads).toBe(firstReads);
    expect(second).toEqual(first);
    // An immutable edit retains the census of every untouched paragraph.
    const updated = {
      ...document,
      package: {
        ...document.package,
        document: {
          ...document.package.document,
          content: [paragraph(0), ...paragraphs.slice(1)],
        },
      },
    };
    allocate(updated, intent);
    expect(unrelatedContentReads).toBe(firstReads);
    for (const mode of [
      { type: "editing", newIds: first.newIds },
      {
        type: "suggesting",
        revision: { id: first.revisionId, author: "Editor" },
        newIds: first.newIds,
      },
    ] as const) {
      const compiled = compileEditorIntent(document, { intent, mode }).unwrap();
      expect(compiled.ops.length).toBeGreaterThan(0);
      for (const op of compiled.ops) {
        if (!("newIds" in op)) continue;
        expect(op.newIds?.revision ?? []).toEqual([]);
        expect(op.newIds?.control ?? []).toEqual([]);
      }
    }
  });

  test("collapsed insertion allocates only endpoint leaves in a long paragraph", () => {
    const source: Paragraph = {
      type: "paragraph",
      paraId: "00000001",
      content: Array.from(
        { length: 2_000 },
        () =>
          ({
            type: "run",
            content: [{ type: "text", text: "x" }],
          }) satisfies Run,
      ),
    };
    const document: Document = { package: { document: { content: [source] } } };
    for (const offset of [0, 1, 1_000, 1_999, 2_000]) {
      const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset };
      const ids = allocateEditorIntentIds(document, {
        type: "replaceText",
        from: at,
        to: at,
        text: "x",
      });
      expect(ids.newIds.revision?.length ?? 0).toBeLessThanOrEqual(24);
      expect(ids.newIds.control?.length ?? 0).toBeLessThanOrEqual(24);
    }
  });

  test("incremental censuses equal fresh censuses through generated immutable versions", () => {
    assertProperty(
      fc.property(
        documentArbitrary,
        fc.array(
          fc.record({
            revision: fc.integer({ min: 1, max: 100 }),
            control: fc.integer({ min: 1, max: 100 }),
          }),
          { minLength: 1, maxLength: 8 },
        ),
        (generated, changes) => {
          const source = generated.package.document.content.find(
            (block) => block.type === "paragraph",
          );
          if (!source) panic("Generated immutable census document has no top-level paragraph.");
          const at = { story: OP_STORIES.MAIN, blockId: source.paraId ?? "", offset: 0 };
          const intent = { type: "replaceText", from: at, to: at, text: "x" } as const;
          const allocate = createEditorIntentIdAllocator();
          let current = generated;
          expect(allocate(current, intent)).toEqual(
            createEditorIntentIdAllocator()(current, intent),
          );
          for (const change of changes) {
            const replacement: Paragraph = {
              ...source,
              content: [
                {
                  type: "insertion",
                  info: { id: change.revision, author: "Editor" },
                  content: [
                    {
                      type: "inlineSdt",
                      properties: { sdtType: "richText", id: change.control },
                      content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
                    },
                  ],
                },
              ],
            };
            current = replaceMainParagraph(current, replacement);
            expect(allocate(current, intent)).toEqual(
              createEditorIntentIdAllocator()(current, intent),
            );
            const candidate = createEditorIntentIdAllocator()(current, intent);
            current = replaceHeaderIdentities(current, candidate);
            expect(allocate(current, intent)).toEqual(
              createEditorIntentIdAllocator()(current, intent),
            );
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  test("generated intent allocations avoid every package identity space", () => {
    const kinds = ["replaceText", "splitParagraph", "joinParagraphs"] as const;
    const exercised = new Set<string>();
    assertProperty(
      fc.property(documentArbitrary, fc.nat(), fc.nat(), (generated, paragraphPick, offsetPick) => {
        const occupied = new Set(packageIdentityKeys(generated.package));
        const firstFree = (space: string) => {
          let id = 1;
          while (occupied.has(`${space}:${id}`)) id += 1;
          return id;
        };
        const outsideRevision = firstFree(IDENTITY_SPACES.REVISION);
        const outsideControl = firstFree(IDENTITY_SPACES.CONTROL);
        const existingParagraphs = new Set(packageParagraphIds(generated.package).map(idKey));
        let outsideParagraph = 1;
        while (existingParagraphs.has(paragraph(outsideParagraph - 1).paraId ?? ""))
          outsideParagraph += 1;
        const document: Document = {
          ...generated,
          package: {
            ...generated.package,
            headers: new Map([
              [
                "rIdHeader",
                {
                  type: "header",
                  hdrFtrType: "default",
                  content: [
                    {
                      type: "paragraph",
                      paraId: paragraph(outsideParagraph - 1).paraId,
                      content: [
                        {
                          type: "insertion",
                          info: { id: outsideRevision, author: "Header" },
                          content: [
                            {
                              type: "inlineSdt",
                              properties: { sdtType: "richText", id: outsideControl },
                              content: [
                                { type: "run", content: [{ type: "text", text: "header" }] },
                              ],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            ]),
          },
        };
        const paragraphs = storyParagraphs(document.package.document).map(
          ({ paragraph: value }) => value,
        );
        const source = paragraphs.at(paragraphPick % paragraphs.length);
        if (!source) panic("Generated allocation document has no paragraph.");
        const next = paragraphs.at((paragraphPick + 1) % paragraphs.length);
        const at = {
          story: OP_STORIES.MAIN,
          blockId: source.paraId ?? "",
          offset: offsetPick % (paragraphLength(source) + 1),
        };
        const intents = [
          { type: "replaceText", from: at, to: at, text: "x" },
          { type: "splitParagraph", at },
          {
            type: "joinParagraphs",
            story: OP_STORIES.MAIN,
            blockId: at.blockId,
            nextBlockId: next?.paraId ?? at.blockId,
          },
        ] as const;
        const identities = new Set(packageIdentityKeys(document.package));
        const paragraphIds = new Set(packageParagraphIds(document.package).map(idKey));
        for (const intent of intents) {
          const ids = allocateEditorIntentIds(document, intent);
          expect(paragraphIds.has(idKey(ids.newBlockId))).toBe(false);
          const revision = [ids.revisionId, ...(ids.newIds.revision ?? [])];
          const control = ids.newIds.control ?? [];
          expect(new Set(revision).size).toBe(revision.length);
          expect(new Set(control).size).toBe(control.length);
          for (const id of revision)
            expect(identities.has(`${IDENTITY_SPACES.REVISION}:${id}`)).toBe(false);
          for (const id of control)
            expect(identities.has(`${IDENTITY_SPACES.CONTROL}:${id}`)).toBe(false);
          exercised.add(intent.type);
        }
      }),
      { numRuns: 100 },
    );
    expect([...exercised].sort()).toEqual([...kinds].sort());
  });
});
