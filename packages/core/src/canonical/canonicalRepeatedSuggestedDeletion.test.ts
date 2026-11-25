/**
 * A suggested deletion of a container's last paragraph records the break
 * before it and its words deleted; the paragraph before it then joins it once
 * accepted. Suggesting that deletion again takes only the paragraph's own
 * words, never the words joined into it, so a batch that edits those words
 * and repeats the deletion leaves what its operations leave one at a time:
 * the edit, and nothing more.
 */

import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type {
  FolioAIEditApplyMode,
  FolioAIEditOperation,
} from "../ai-edits/types";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const MODES = [
  "suggested",
  "tracked-changes",
] as const satisfies readonly FolioAIEditApplyMode[];

/** Edits of the paragraph before the deleted one, with what each leaves. */
const EDITS = [
  { find: "Launch ", replace: "revised", leaves: "revised🚀" },
  { find: "Launch ", replace: "Go ", leaves: "Go 🚀" },
  {
    find: "Launch ",
    replace: "Lift-off at once ",
    leaves: "Lift-off at once 🚀",
  },
  { find: "🚀", replace: "now", leaves: "Launch now" },
] as const;

const open = async () => {
  const input = (
    await prepareCanonicalDocxInput(
      await createDocx(
        fromMarkdown(
          "# Emoji\n\nApproved 👍🏽 by the family 👨‍👩‍👧 today.\n\nLaunch 🚀\n\n🎉 opens the party.",
        ),
      ),
    )
  )
    .unwrap()
    .slice().buffer;
  const reviewer = await FolioDocxReviewer.fromBuffer(input, {
    author: "Test",
    session: "canonical",
  });
  const [, , launch, last] = reviewer.snapshot().blocks;
  if (launch?.text !== "Launch 🚀" || last?.text !== "🎉 opens the party.")
    throw new Error("Missing fixture");
  return { reviewer, launch, last };
};

type Fixture = Awaited<ReturnType<typeof open>>;

/**
 * The blocks, and the text each kind of change holds in each paragraph: one
 * batch stamps its records alike, so adjacent ones read as one change where
 * separate calls leave two.
 */
const outcome = (reviewer: FolioDocxReviewer) => {
  const changes = new Map<string, string>();
  for (const { type, text, blockId } of reviewer.getChanges()) {
    const key = `${type} ${String(blockId)}`;
    changes.set(key, (changes.get(key) ?? "") + text);
  }
  return {
    blocks: reviewer.snapshot().blocks.map(({ id, text }) => ({ id, text })),
    changes: Object.fromEntries(changes),
  };
};

for (const mode of MODES)
  for (const { find, replace, leaves } of EDITS)
    test(`${mode}: replacing "${find}" with "${replace}" beside a repeated deletion of the last paragraph equals one at a time`, async () => {
      const batchOf = ({ launch, last }: Fixture): FolioAIEditOperation[] => [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: launch.id,
          find,
          replace,
        },
        { id: "again", type: "deleteBlock", blockId: last.id },
      ];
      const batched = await open();
      const sequential = await open();
      for (const { reviewer, last } of [batched, sequential])
        expect(
          reviewer.applyOperations(
            [{ id: "delete", type: "deleteBlock", blockId: last.id }],
            { mode, undefinedReferences: "refuse" },
          ).skipped,
        ).toEqual([]);
      const batch = batched.reviewer.applyOperations(batchOf(batched), {
        mode,
        undefinedReferences: "refuse",
      });
      const skipped = batchOf(sequential).flatMap(
        (operation) =>
          sequential.reviewer.applyOperations([operation], {
            mode,
            undefinedReferences: "refuse",
          }).skipped,
      );
      expect(batch.skipped).toEqual(skipped);
      expect(outcome(batched.reviewer)).toEqual(outcome(sequential.reviewer));
      if (mode === "suggested")
        expect(
          batched.reviewer
            .snapshot()
            .blocks.find(({ id }) => id === batched.launch.id)?.text,
        ).toBe(leaves);
    });
