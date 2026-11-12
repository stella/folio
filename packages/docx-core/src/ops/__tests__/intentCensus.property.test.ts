/**
 * The ids an editor intent reads as taken, from the package's kept indexes,
 * are those a census of each version read afresh holds.
 *
 * The indexes count a new version along the paths the edit changed:
 * paragraphs, tables and content controls alike. A wrong count there would
 * let a later edit mint an id the document already holds, so every version
 * of a random run of edits, direct and tracked, in the body and in notes,
 * through tables and stacked row controls, and back to earlier versions, is
 * read both ways.
 */
import { describe, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { packageIdOccupancy } from "../editorIntentCensus";
import {
  IDENTITY_SPACES,
  idKey,
  packageIdentityKeys,
  packageParagraphIds,
  packageReservedIdentityKeys,
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
import { type CensusIds, expectOccupancyOf } from "./occupancyProbe";

setDefaultTimeout(propertyTestTimeout(240_000));

const STACKED_ROW_PARAGRAPHS = ["7AAAAAA1", "7AAAAAA2"] as const;

/** A table whose two rows sit in one content control, recorded on each row. */
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

/** The ids of each space a census of the whole package holds. */
const censusOf = (document: Document): CensusIds => {
  const keys = [
    ...packageIdentityKeys(document.package),
    ...packageReservedIdentityKeys(document.package),
  ];
  const idsIn = (space: string) =>
    new Set(
      keys
        .filter((key) => key.startsWith(`${space}:`))
        .map((key) => Number(key.slice(space.length + 1))),
    );
  return {
    revisions: idsIn(IDENTITY_SPACES.REVISION),
    controls: idsIn(IDENTITY_SPACES.CONTROL),
    paragraphs: new Set(packageParagraphIds(document.package).map(idKey)),
  };
};

const expectCensusOf = (document: Document): void =>
  expectOccupancyOf(packageIdOccupancy(document.package), censusOf(document));

describe("the id census an editor intent reads", () => {
  test("equals a census read afresh over a run of edits", async () => {
    await assertProperty(
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        fc.boolean(),
        fc.array(
          fc.record({
            seed: opSeedArbitrary,
            tracked: fc.boolean(),
            story: fc.nat(),
            // Typing into or splitting a paragraph of one of the stacked
            // rows, whose control counts once.
            row: fc.constantFrom(undefined, 0, 1),
            split: fc.boolean(),
            // Reading an earlier version again, as undo brings one back.
            back: fc.option(fc.nat(), { freq: 4 }),
          }),
          { minLength: 1, maxLength: 12 },
        ),
        (generated, stacked, edits) => {
          let document = stacked ? withStackedRowControl(generated) : generated;
          const versions = [document];
          expectCensusOf(document);
          edits.forEach(
            ({ seed, tracked, story: at, row, split, back }, index) => {
              if (back !== null) {
                document = versions[back % versions.length] ?? document;
                expectCensusOf(document);
                return;
              }
              const stories = documentStories(document);
              const story = stories[at % stories.length] ?? OP_STORIES.MAIN;
              const target =
                stacked && row !== undefined
                  ? STACKED_ROW_PARAGRAPHS[row]
                  : undefined;
              const drawn = (): DocumentOp =>
                tracked && story === OP_STORIES.MAIN
                  ? trackedOpFor(document, seed, index)
                  : opForStory({ document, seed, story });
              const inRow = (blockId: string): DocumentOp =>
                split
                  ? {
                      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
                      at: { story: OP_STORIES.MAIN, blockId, offset: 0 },
                      newBlockId: (0x7b000000 + index)
                        .toString(16)
                        .toUpperCase(),
                    }
                  : {
                      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                      at: { story: OP_STORIES.MAIN, blockId, offset: 0 },
                      text: "x",
                      runProps: INHERIT_RUN_PROPS,
                    };
              const op = target === undefined ? drawn() : inRow(target);
              const applied = applyDocumentOp(document, op);
              if (applied.isErr()) return;
              document = applied.value.document;
              versions.push(document);
              expectCensusOf(document);
            },
          );
        },
      ),
      {
        id: "editor intent census kept across versions equals a fresh census",
        numRuns: 200,
      },
    );
  });
});
