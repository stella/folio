/**
 * A census of a document that shares records with documents already counted
 * equals a census walking every record.
 *
 * The id census keeps what each unchanged record contributes, so a document
 * an edit made is counted by walking only what the edit replaced. A wrong
 * count there would let a later edit mint an id the document already holds,
 * so every version of a random run of edits, direct and tracked, in the body
 * and in notes, through tables and stacked row controls, is counted both ways.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  identityKeysIn,
  packageIdentityKeys,
  packageReservedIdentityKeys,
  reservedIdentityKeysIn,
} from "../ids";
import { documentStories } from "../stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "../types";
import {
  documentArbitrary,
  opForStory,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

const STACKED_ROW_PARAGRAPHS = ["7AAAAAA1", "7AAAAAA2"] as const;

/** One content control over two rows, recorded on each row. */
const withStackedRowControl = (document: Document): Document => {
  const control = { sdtType: "richText" as const, id: 4242, tag: "rows" };
  const row = (paraId: string) => ({
    type: "tableRow" as const,
    cells: [
      {
        type: "tableCell" as const,
        content: [
          {
            type: "paragraph" as const,
            paraId,
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: "cell" }],
              },
            ],
          },
        ],
      },
    ],
    contentControls: [control],
  });
  return {
    ...document,
    package: {
      ...document.package,
      document: {
        ...document.package.document,
        content: [
          ...document.package.document.content,
          {
            type: "table",
            columnWidths: [4800],
            rows: STACKED_ROW_PARAGRAPHS.map(row),
          },
        ],
      },
    },
  };
};

/** The package census as edits read it, reusing what earlier versions counted. */
const counted = (document: Document) => ({
  package: [...packageIdentityKeys(document.package)],
  packageReserved: [...packageReservedIdentityKeys(document.package)],
});

/** The same census walking every record, as if nothing were counted before. */
const countedInFull = (document: Document) => {
  const stories = {
    ...document.package,
    document: { ...document.package.document, sections: undefined },
  };
  return {
    package: identityKeysIn(stories),
    // Reserved ids are read off the whole package, its section view included.
    packageReserved: reservedIdentityKeysIn(document.package),
  };
};

describe("the id census of a shared document", () => {
  test("equals a census walking every record over a run of edits", async () => {
    await assertProperty(
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        fc.boolean(),
        fc.array(
          fc.record({
            seed: opSeedArbitrary,
            tracked: fc.boolean(),
            story: fc.nat(),
            // Typing into one of the stacked rows, whose control the census
            // counts once across both.
            row: fc.constantFrom(undefined, 0, 1),
          }),
          { minLength: 1, maxLength: 12 },
        ),
        (generated, stacked, edits) => {
          let document = stacked ? withStackedRowControl(generated) : generated;
          expect(counted(document)).toEqual(countedInFull(document));
          edits.forEach(({ seed, tracked, story: storyIndex, row }, index) => {
            const stories = documentStories(document);
            const story =
              stories[storyIndex % stories.length] ?? OP_STORIES.MAIN;
            const target =
              stacked && row !== undefined
                ? STACKED_ROW_PARAGRAPHS[row]
                : undefined;
            const drawn = (): DocumentOp =>
              tracked && story === OP_STORIES.MAIN
                ? trackedOpFor(document, seed, index)
                : opForStory({ document, seed, story });
            const op: DocumentOp =
              target === undefined
                ? drawn()
                : {
                    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                    at: {
                      story: OP_STORIES.MAIN,
                      blockId: target,
                      offset: 0,
                    },
                    text: "x",
                    runProps: INHERIT_RUN_PROPS,
                  };
            const applied = applyDocumentOp(document, op);
            if (applied.isErr()) return;
            document = applied.value.document;
            expect(counted(document)).toEqual(countedInFull(document));
          });
        },
      ),
      { id: "identity census shares records soundly", numRuns: 200 },
    );
  });

  test("reads content a caller changes between two censuses as it is now", () => {
    const content: Record<string, unknown>[] = [
      { type: "run", content: [{ type: "text", text: "pending" }] },
    ];
    const paragraph = { type: "paragraph", content };
    expect(identityKeysIn(paragraph)).toEqual([]);
    expect(reservedIdentityKeysIn(paragraph)).toEqual([]);
    content.push({
      type: "insertion",
      info: { id: 41, author: "Reviewer" },
      content: [],
      resolutionJoins: {
        retainedAfter: [
          { source: [{ space: "revision", id: 42 }], target: [] },
        ],
      },
    });
    expect(identityKeysIn(paragraph)).toEqual(["revision:41"]);
    expect(reservedIdentityKeysIn(paragraph)).toEqual(["revision:42"]);
  });
});
