/**
 * A package's key index, derived from the package indexed before it, counts
 * what a census walking every record counts.
 *
 * Each index is found from the last one by walking only what an edit
 * replaced, so a wrong change there would let a later edit mint an id the
 * document already holds. Every version of a random run of edits, direct
 * and tracked, in the body and in notes, through tables and stacked row
 * controls, is indexed and compared with the full census, including
 * versions branching from an earlier one and versions indexed out of order.
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
  countKeys,
  type IdentityKeyIndex,
  identityKeysIn,
  packageIdentityKeyIndex,
  packageReservedKeyIndex,
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

/** The census walking every record, as if nothing were counted before. */
const countedInFull = (document: Document) => {
  const stories = {
    ...document.package,
    document: { ...document.package.document, sections: undefined },
  };
  return {
    identity: countKeys(identityKeysIn(stories)),
    // Reserved ids are read off the whole package, its section view included.
    reserved: countKeys(reservedIdentityKeysIn(document.package)),
  };
};

/**
 * The index agrees with `expected` on every key in `asked`, through its
 * lookups and then, when `listed`, through its full list.
 */
const expectIndexed = ({
  index,
  expected,
  asked,
  listed,
}: {
  index: IdentityKeyIndex;
  expected: ReadonlyMap<string, number>;
  asked: ReadonlySet<string>;
  listed: boolean;
}): void => {
  for (const key of asked) {
    expect([key, index.count(key)]).toEqual([key, expected.get(key) ?? 0]);
    expect(index.has(key)).toBe((expected.get(key) ?? 0) > 0);
  }
  if (listed) expect(new Map(index.counts())).toEqual(new Map(expected));
};

describe("the key index of a package", () => {
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
            // Edit an earlier version instead of the latest one.
            branch: fc.option(fc.nat(), { freq: 4 }),
            // Index the edited version only after the next one.
            deferred: fc.boolean(),
            listed: fc.boolean(),
          }),
          { minLength: 1, maxLength: 40 },
        ),
        (generated, stacked, edits) => {
          const start = stacked ? withStackedRowControl(generated) : generated;
          const versions: Document[] = [start];
          const asked = new Set<string>(["revision:0", "control:0"]);
          const check = (document: Document, listed: boolean): void => {
            const expected = countedInFull(document);
            for (const key of expected.identity.keys()) asked.add(key);
            for (const key of expected.reserved.keys()) asked.add(key);
            expectIndexed({
              index: packageIdentityKeyIndex(document.package),
              expected: expected.identity,
              asked,
              listed,
            });
            expectIndexed({
              index: packageReservedKeyIndex(document.package),
              expected: expected.reserved,
              asked,
              listed,
            });
          };
          check(start, false);
          let pending: Document | undefined;
          edits.forEach((edit, index) => {
            const { seed, tracked, story: storyIndex, row, branch } = edit;
            const latest = versions.at(-1) ?? start;
            const document =
              branch === null
                ? latest
                : (versions[branch % versions.length] ?? latest);
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
            const next = applied.value.document;
            versions.push(next);
            if (edit.deferred && pending === undefined) {
              pending = next;
              return;
            }
            check(next, edit.listed);
            if (pending !== undefined) check(pending, edit.listed);
            pending = undefined;
          });
          if (pending !== undefined) check(pending, true);
          // Every version again: settling the later indexes left the
          // earlier ones counting what they counted.
          for (const version of versions) check(version, false);
          for (const version of versions) check(version, true);
        },
      ),
      { id: "identity key index follows edits soundly", numRuns: 120 },
    );
  });
});

describe("the key index after the newest indexed package is gone", () => {
  // An edit's intermediate version is indexed and then dropped; once it is
  // collected, the next version is found from an older one still held.
  test("still equals a census walking every record", async () => {
    await assertProperty(
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        fc.array(
          fc.record({
            kept: opSeedArbitrary,
            dropped: opSeedArbitrary,
            tracked: fc.boolean(),
            collect: fc.boolean(),
          }),
          { minLength: 1, maxLength: 12 },
        ),
        (start, edits) => {
          let latest = start;
          const check = (document: Document): void => {
            const expected = countedInFull(document);
            const asked = new Set([
              ...expected.identity.keys(),
              ...expected.reserved.keys(),
            ]);
            expectIndexed({
              index: packageIdentityKeyIndex(document.package),
              expected: expected.identity,
              asked,
              listed: true,
            });
            expectIndexed({
              index: packageReservedKeyIndex(document.package),
              expected: expected.reserved,
              asked,
              listed: true,
            });
          };
          check(latest);
          edits.forEach(({ kept, dropped, tracked, collect }, index) => {
            const opOf = (document: Document, seed: typeof kept) =>
              tracked
                ? trackedOpFor(document, seed, index)
                : opForStory({ document, seed, story: OP_STORIES.MAIN });
            // Indexed on the way, then let go.
            (() => {
              const transient = applyDocumentOp(latest, opOf(latest, dropped));
              if (transient.isOk())
                packageIdentityKeyIndex(transient.value.document.package);
            })();
            if (collect) Bun.gc(true);
            const next = applyDocumentOp(latest, opOf(latest, kept));
            if (next.isErr()) return;
            latest = next.value.document;
            check(latest);
          });
        },
      ),
      { numRuns: 60, id: "the key index after the newest package is gone" },
    );
  });
});

describe("the key index across documents", () => {
  // Several documents edited in turn, each edit's intermediate version
  // dropped and sometimes collected: a version is often found from another
  // document's package, which must count what a full census counts.
  test("equals a census walking every record whichever package it derives from", async () => {
    await assertProperty(
      fc.property(
        fc.array(fc.oneof(documentArbitrary, reviewDocumentArbitrary), {
          minLength: 2,
          maxLength: 4,
        }),
        fc.array(
          fc.record({
            document: fc.nat(),
            kept: opSeedArbitrary,
            dropped: opSeedArbitrary,
            tracked: fc.boolean(),
            collect: fc.boolean(),
            // Edit an earlier version of the document instead of its latest.
            branch: fc.option(fc.nat(), { freq: 3 }),
          }),
          { minLength: 2, maxLength: 16 },
        ),
        (generated, edits) => {
          // Every document but the first holds a header of its own, under
          // its own relationship id, with a record in it, so a version found
          // from another document's package counts headers in and out.
          const documents = generated.map((document, index): Document =>
            index === 0
              ? document
              : {
                  ...document,
                  package: {
                    ...document.package,
                    headers: new Map([
                      [
                        `rIdIndexed${index}`,
                        {
                          type: "header",
                          hdrFtrType: "default",
                          content: [
                            {
                              type: "paragraph",
                              paraId: (0x7b000000 + index)
                                .toString(16)
                                .toUpperCase(),
                              content: [
                                {
                                  type: "insertion",
                                  info: { id: 9000 + index, author: "Header" },
                                  content: [
                                    {
                                      type: "run",
                                      content: [{ type: "text", text: "h" }],
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
                },
          );
          const histories = documents.map((document) => [document]);
          const check = (document: Document): void => {
            const expected = countedInFull(document);
            const asked = new Set([
              ...expected.identity.keys(),
              ...expected.reserved.keys(),
            ]);
            expectIndexed({
              index: packageIdentityKeyIndex(document.package),
              expected: expected.identity,
              asked,
              listed: true,
            });
            expectIndexed({
              index: packageReservedKeyIndex(document.package),
              expected: expected.reserved,
              asked,
              listed: true,
            });
          };
          for (const document of documents) check(document);
          edits.forEach((edit, index) => {
            const history = histories[edit.document % histories.length] ?? [];
            const latest = history.at(-1);
            if (latest === undefined) return;
            const base =
              edit.branch === null
                ? latest
                : (history[edit.branch % history.length] ?? latest);
            const opOf = (document: Document, seed: typeof edit.kept) =>
              edit.tracked
                ? trackedOpFor(document, seed, index)
                : opForStory({ document, seed, story: OP_STORIES.MAIN });
            (() => {
              const transient = applyDocumentOp(base, opOf(base, edit.dropped));
              if (transient.isOk())
                packageIdentityKeyIndex(transient.value.document.package);
            })();
            if (edit.collect) Bun.gc(true);
            const next = applyDocumentOp(base, opOf(base, edit.kept));
            if (next.isErr()) return;
            history.push(next.value.document);
            check(next.value.document);
          });
        },
      ),
      { numRuns: 60, id: "the key index across documents" },
    );
  });
});
