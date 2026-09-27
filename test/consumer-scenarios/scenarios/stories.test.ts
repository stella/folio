/**
 * Operations aimed where edits collide, deterministically, so the coverage
 * ledger's required cells (coverage-expectations.json) do not hang on a
 * seed: every operation type in each header, footer and note of the stories
 * fixture, every target feature of its body, and a document carried across
 * sessions (saved, reopened, then reopened by another reviewer). Each
 * applied operation must do what it asked (support/oracle.ts), and the
 * document must still save, reopen and read alike, its secondary stories
 * included.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import type { FolioDocumentStoryHandle } from "@stll/folio-core/server";

import type { StepKind } from "../support/coverage.ts";
import { openReviewer, storiesDocument } from "../support/documents.ts";
import { SECOND_REVIEWER } from "../support/fuzz.ts";
import { assertHealthy } from "../support/invariants.ts";
import { assertRequestedOutcome, capture } from "../support/oracle.ts";
import {
  type Block,
  coreBatch,
  GENERATORS,
  MODES,
  type Mode,
  type Operation,
  supports,
} from "../support/operations.ts";
import { createRandom, type Random } from "../support/random.ts";
import {
  biasedPicker,
  blocksOfStory,
  type Feature,
  FEATURES,
  featureIndex,
  type Picker,
} from "../support/targets.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
const MAIN: FolioDocumentStoryHandle = { type: "main" };

/** Apply `operation` to `story` as a one-operation core batch, and check it. */
const applyChecked = async (
  reviewer: Reviewer,
  story: FolioDocumentStoryHandle,
  operation: Operation,
  mode: Mode,
  step: StepKind,
): Promise<boolean> => {
  const pre = await capture(reviewer, mode, { story, step });
  const batch = coreBatch([operation], mode);
  const result =
    story.type === "main"
      ? reviewer.applyDocumentOperations(batch as never)
      : reviewer.applyDocumentOperationsToStory({ story, batch: batch as never });
  const applied = result.applied.length > 0;
  await assertRequestedOutcome(
    reviewer,
    pre,
    { applied: applied ? batch.operations : [], attempted: batch.operations },
    `${JSON.stringify(story)} / ${mode}: ${JSON.stringify(operation)}`,
  );
  return applied;
};

/** The pieces of the stories fixture every sweep reads. */
const secondaryStories = (reviewer: Reviewer) =>
  reviewer
    .listStories()
    .map((story) => story.handle)
    .filter((handle) => handle.type !== "main");

const storyTexts = (reviewer: Reviewer) =>
  secondaryStories(reviewer).map((handle) =>
    blocksOfStory(reviewer, handle).map((block) => block.text),
  );

/** Saves, reopens as shown, and every header, footer and note reopens as shown. */
const assertStoriesHealthy = async (reviewer: Reviewer, mode: Mode, context: string) => {
  const suggested = mode === "suggested";
  const { reopened } = await assertHealthy(
    reviewer,
    context,
    suggested && reviewer.getChanges().length > 0 ? { compare: false } : {},
  );
  if (!suggested) {
    assert.deepEqual(
      storyTexts(reopened),
      storyTexts(reviewer),
      `${context}: a story changed on save`,
    );
  }
};

/** A picker aimed at blocks with `feature`, and at boundary spans in them. */
const featurePicker = (reviewer: Reviewer, feature: Feature, random: Random): Picker => {
  const index = featureIndex(reviewer);
  const aimed = biasedPicker(random, { index, recent: [], uniform: 0 });
  return {
    ...aimed,
    block: (candidates) => {
      const hot = candidates.filter((block) => index.features.get(block.id)?.has(feature));
      return random.pick(hot.length > 0 ? hot : candidates);
    },
  };
};

/** The operation types a feature sweep tries on each feature. */
const FEATURE_OPERATIONS = [
  "replaceInBlock",
  "replaceRange",
  "formatRange",
  "commentOnRange",
  "insertAfterBlock",
  "splitBlock",
  "setBlockParagraphProperties",
  "replaceBlock",
] as const;

describe("operations in every header, footer and note", () => {
  for (const mode of MODES) {
    test(`stories / ${mode}: every operation type in every secondary story does what it asked`, async () => {
      const reviewer = await openReviewer(await storiesDocument());
      const random = createRandom(mode.length * 97);
      const applied = new Map<string, number>();
      // One of each kind; the fuzz flows reach the first-page and even-page parts.
      const stories = secondaryStories(reviewer).filter(
        (story, index, all) => all.findIndex((other) => other.type === story.type) === index,
      );
      for (const story of stories) {
        for (const type of Object.keys(GENERATORS)) {
          if (!supports(type, mode)) continue;
          const blocks = blocksOfStory(reviewer, story) as Block[];
          const pick = biasedPicker(random, { index: featureIndex(reviewer, story), recent: [] });
          const operation = GENERATORS[type]?.(blocks, random, pick);
          if (!operation) continue;
          if (await applyChecked(reviewer, story, operation, mode, "fresh")) {
            applied.set(story.type, (applied.get(story.type) ?? 0) + 1);
          }
        }
        await assertStoriesHealthy(reviewer, mode, `${JSON.stringify(story)} / ${mode}`);
      }
      for (const kind of ["header", "footer", "footnote", "endnote"]) {
        assert.ok((applied.get(kind) ?? 0) > 0, `no operation applied in a ${kind}`);
      }
    });
  }
});

describe("operations aimed at every feature of the body", () => {
  for (const mode of MODES) {
    test(`stories / ${mode}: operations at each feature do what they asked`, async () => {
      const reviewer = await openReviewer(await storiesDocument());
      // A pending revision to aim at, whatever the mode.
      const payment = reviewer.getContent().find((block) => block.text.startsWith("Payment is"));
      assert.ok(payment);
      await applyChecked(
        reviewer,
        MAIN,
        { type: "replaceInBlock", blockId: payment.id, find: "receipt", replace: "invoice" },
        "tracked-changes",
        "fresh",
      );
      const random = createRandom(mode.length * 131);
      for (const [index, feature] of FEATURES.entries()) {
        // Two operation types per feature, rotating, so each type meets several features.
        for (const offset of [0, 3]) {
          const type = FEATURE_OPERATIONS[(index + offset) % FEATURE_OPERATIONS.length] as string;
          if (!supports(type, mode)) continue;
          const operation = GENERATORS[type]?.(
            reviewer.getContent() as Block[],
            random,
            featurePicker(reviewer, feature, random),
          );
          if (operation) await applyChecked(reviewer, MAIN, operation, mode, "fresh");
        }
      }
      await assertStoriesHealthy(reviewer, mode, `features / ${mode}`);
      // The text box survives every edit around its paragraph (none deletes it).
      assert.ok(
        JSON.stringify(
          (await openReviewer(new Uint8Array(await reviewer.toBuffer()))).toDocument().package
            .document.content,
        ).includes('"shapeType":"textBox"'),
        `features / ${mode}: the text box is gone`,
      );
    });

    // A reader lists a text box's paragraphs with the body. A fresh document:
    // with a change pending in the paragraph that draws the box too, resolving
    // after a reopen writes malformed XML (TEXT_BOX_RESOLVE_MALFORMED_XML).
    test(`stories / ${mode}: operations in a text box do what they asked`, async () => {
      const reviewer = await openReviewer(await storiesDocument());
      const random = createRandom(mode.length * 151);
      for (const type of FEATURE_OPERATIONS) {
        const index = featureIndex(reviewer);
        const blocks = (reviewer.getContent() as Block[]).filter((block) =>
          index.inTextBox.has(block.id),
        );
        if (!supports(type, mode) || blocks.length === 0) continue;
        const pick = biasedPicker(random, { index, recent: [] });
        const operation = GENERATORS[type]?.(blocks, random, pick);
        if (operation) await applyChecked(reviewer, MAIN, operation, mode, "fresh");
      }
      await assertStoriesHealthy(reviewer, mode, `text box / ${mode}`);
      assert.ok(
        reviewer.getContent().some((block) => featureIndex(reviewer).inTextBox.has(block.id)),
        `text box / ${mode}: no paragraph is left in the text box`,
      );
    });
  }
});

describe("a document carried across sessions", () => {
  for (const mode of MODES) {
    test(`stories / ${mode}: edits after a reopen and by another reviewer do what they asked`, async () => {
      let reviewer = await openReviewer(await storiesDocument());
      const random = createRandom(mode.length * 173);
      const sessions: [StepKind, (bytes: Uint8Array) => Promise<Reviewer>][] = [
        ["fresh", async () => reviewer],
        ["reopened", (bytes) => openReviewer(bytes)],
        ["newReviewer", (bytes) => openReviewer(bytes, SECOND_REVIEWER)],
      ];
      for (const [step, open] of sessions) {
        const { bytes } = await assertHealthy(
          reviewer,
          `${mode} before ${step}`,
          mode === "suggested" && reviewer.getChanges().length > 0 ? { compare: false } : {},
        );
        reviewer = await open(bytes);
        for (const story of [MAIN, ...secondaryStories(reviewer).slice(0, 1)]) {
          for (const type of ["replaceInBlock", "insertAfterBlock", "formatRange"]) {
            if (!supports(type, mode)) continue;
            const index = featureIndex(reviewer, story);
            const blocks = (blocksOfStory(reviewer, story) as Block[]).filter(
              (block) => !index.features.get(block.id)?.has("commentAnchor"),
            );
            const pick = biasedPicker(random, { index, recent: [] });
            const operation = GENERATORS[type]?.(blocks, random, pick);
            if (operation) await applyChecked(reviewer, story, operation, mode, step);
          }
        }
        // The same reviewer keeps working after a selective save.
        const selective = await reviewer.save({ repack: "refuse" });
        if (selective.type === "selective") {
          await openReviewer(new Uint8Array(selective.buffer));
        }
      }
      await assertStoriesHealthy(reviewer, mode, `sessions / ${mode}`);
    });
  }
});
