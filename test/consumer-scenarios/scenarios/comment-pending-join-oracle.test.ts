/**
 * A comment on either paragraph of a pending join, added before or after the
 * join, then each structural edit of the join's source, its destination or
 * the paragraph after them. The oracle checks every step: a comment goes once
 * accepted exactly when the words it covers do.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { hashFolioAIBlockText } from "@stll/folio-core/server";

import { parseFlowFile } from "../support/flow-file.ts";
import { runFlowFile, vacuousSteps } from "../support/fuzz.ts";

const BLOCKS = [
  { id: "412DDC7F", text: "Terms" },
  { id: "3CEFFC83", text: "Numbered clause one" },
  { id: "31617F9A", text: "Unnumbered body text." },
  { id: "5E88C024", text: "Numbered clause two" },
] as const;
type Block = (typeof BLOCKS)[number];

const batch = (seed: number, operations: readonly object[]) => ({
  action: "core batch",
  seed,
  atomic: false,
  operations,
});

const comment = (block: Block) =>
  batch(2, [
    {
      type: "commentOnRange",
      range: {
        type: "textRange",
        story: "main",
        blockId: block.id,
        startOffset: 0,
        endOffset: block.text.length,
        selectedTextHash: hashFolioAIBlockText(block.text),
      },
      comment: { text: "Check the joined words." },
    },
  ]);

const edits = (
  source: Block,
  destination: Block,
  next: Block | undefined,
): [string, object[]][] => [
  ["delete the source", [{ type: "deleteBlock", blockId: source.id }]],
  [
    "delete both",
    [
      { type: "deleteBlock", blockId: source.id },
      { type: "deleteBlock", blockId: destination.id },
    ],
  ],
  [
    "replace the destination",
    [{ type: "replaceBlock", blockId: destination.id, text: "New words." }],
  ],
  [
    "replace the source",
    [{ type: "replaceBlock", blockId: source.id, text: "New words." }],
  ],
  [
    "split the destination",
    [{ type: "splitBlock", blockId: destination.id, offset: 3 }],
  ],
  ...(next === undefined
    ? []
    : ([
        // The story's last paragraph keeps its mark, so only a destination
        // with a paragraph after it can be deleted on its own.
        [
          "delete the destination",
          [{ type: "deleteBlock", blockId: destination.id }],
        ],
        [
          "join the destination on",
          [
            {
              type: "mergeBlockWithNext",
              blockId: destination.id,
              separator: " ",
            },
          ],
        ],
        [
          "delete the paragraph after",
          [{ type: "deleteBlock", blockId: next.id }],
        ],
        [
          "delete the destination and the paragraph after",
          [
            { type: "deleteBlock", blockId: next.id },
            { type: "deleteBlock", blockId: destination.id },
          ],
        ],
      ] as [string, object[]][])),
];

// A join runs into the next paragraph, or over one pending deletion into
// the paragraph after it.
for (const over of [0, 1])
  for (let at = 0; at + over + 1 < BLOCKS.length; at++) {
    const source = BLOCKS[at];
    const between = over === 1 ? BLOCKS[at + 1] : undefined;
    const destination = BLOCKS[at + over + 1];
    assert.ok(source !== undefined && destination !== undefined);
    const next = BLOCKS[at + over + 2];
    const route = between === undefined ? "" : ` over ${between.text}`;
    for (const [name, operations] of edits(source, destination, next))
      for (const on of [source, destination])
        for (const order of ["before", "after"] as const)
          test(`join ${source.text}${route} into ${destination.text}, comment on ${on.text} ${order}, ${name}`, async () => {
            const join = batch(1, [
              {
                type: "mergeBlockWithNext",
                blockId: source.id,
                separator: " ",
              },
            ]);
            const steps = [
              ...(between === undefined
                ? []
                : [batch(4, [{ type: "deleteBlock", blockId: between.id }])]),
              ...(order === "before"
                ? [comment(on), join]
                : [join, comment(on)]),
              batch(3, operations),
            ];
            const run = await runFlowFile(
              parseFlowFile({
                version: 1,
                kind: "random",
                generation: "targeted",
                fixture: "directNumbered",
                mode: "tracked-changes",
                seed: 1,
                steps,
              }),
            );
            assert.deepEqual(vacuousSteps(run), []);
          });
  }
