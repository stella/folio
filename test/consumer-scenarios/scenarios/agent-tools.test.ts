/**
 * The `@stll/folio-agents` tools over a reviewer bridge, the way a model
 * drives them: read, find, suggest, comment, reply, resolve; then the
 * mistakes models make. Every call either refuses (with an error or an
 * issue) or leaves a document that saves, reopens and reads alike.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";

import { FIXTURE_NAMES, FIXTURES, openReviewer } from "../support/documents.ts";
import { assertHealthy, visibleState } from "../support/invariants.ts";
import { type Block, MISTAKES, MODES, type Mode } from "../support/operations.ts";
import { createRandom } from "../support/random.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Row = { blockId: string; kind: string; text: string; blockTextHash: string };
type SuggestResult = {
  applied: unknown[];
  skipped: unknown[];
  issues: { code: string }[];
};

const call = (bridge: ReturnType<typeof createReviewerBridge>, name: string, args: unknown) =>
  executeFolioToolCallUntyped(name, args, bridge, {});

const ok = <T>(result: ReturnType<typeof call>, what: string): T => {
  if (!result.ok) {
    assert.fail(`${what} failed: ${result.error}`);
  }
  return result.result as T;
};

const persistedFor = (reviewer: Reviewer, mode: Mode) =>
  mode === "suggested" ? { persisted: visibleState(reviewer) } : {};

describe("agent tools", () => {
  for (const name of FIXTURE_NAMES) {
    for (const mode of MODES) {
      test(`${name} / ${mode}: a model's read → find → suggest → comment session saves`, async () => {
        const reviewer = await openReviewer(await FIXTURES[name]());
        const persisted = persistedFor(reviewer, mode);
        const bridge = createReviewerBridge(reviewer, { mode });

        const rows = ok<Row[]>(call(bridge, "read_document", {}), "read_document");
        assert.ok(rows.length > 0);
        const outline = ok<{ sections: { handle: unknown }[] }>(
          call(bridge, "get_document_outline", {}),
          "get_document_outline",
        );
        const [firstSection] = outline.sections;
        if (firstSection) {
          ok(call(bridge, "read_section", { handle: firstSection.handle }), "read_section");
        }
        for (const story of ok<{ handle: unknown }[]>(
          call(bridge, "list_stories", {}),
          "list_stories",
        )) {
          ok(call(bridge, "read_story", { handle: story.handle }), "read_story");
        }

        const target = rows.find((row) => row.kind === "paragraph" && /\w{4,}/u.test(row.text));
        assert.ok(target, "no paragraph to edit");
        const word = /\w{4,}/u.exec(target.text)?.[0] ?? "";
        const found = ok<{ matches: { range: unknown }[] }>(
          call(bridge, "find_text", { query: word, wholeWord: true }),
          "find_text",
        );
        assert.ok(found.matches.length > 0, `find_text found no "${word}"`);

        const suggested = ok<SuggestResult>(
          call(bridge, "suggest_changes", {
            operations: [
              { type: "replaceRange", range: found.matches[0]?.range, replace: "revised" },
              {
                type: "insertAfterBlock",
                blockId: target.blockId,
                text: "An added sentence.",
                precondition: { blockTextHash: target.blockTextHash },
              },
            ],
          }),
          "suggest_changes",
        );
        assert.equal(suggested.applied.length + suggested.skipped.length, 2);
        await assertHealthy(reviewer, `${name} / ${mode} after suggest_changes`, persisted);

        const comment = call(bridge, "add_comment", {
          blockId: target.blockId,
          quote: word,
          text: "Please confirm.",
        });
        if (mode === "suggested") {
          // Comments are not suggestions; the tool must say so, not half-apply.
          assert.ok(
            !comment.ok ||
              (comment.result as SuggestResult).applied.length === 0 ||
              reviewer.getComments().length > 0,
          );
        } else {
          ok(comment, "add_comment");
        }
        const comments = ok<{ id: string; resolved: boolean }[]>(
          call(bridge, "read_comments", {}),
          "read_comments",
        );
        const [first] = comments;
        if (first) {
          ok(
            call(bridge, "reply_comment", { commentId: first.id, text: "Confirmed." }),
            "reply_comment",
          );
          ok(call(bridge, "resolve_comment", { commentId: first.id }), "resolve_comment");
        }
        ok(call(bridge, "read_changes", {}), "read_changes");
        // Comments are not suggestions: they reach the package in every mode.
        const afterComments =
          persisted.persisted === undefined
            ? {}
            : { persisted: { ...persisted.persisted, comments: visibleState(reviewer).comments } };
        await assertHealthy(reviewer, `${name} / ${mode} after the comment round`, afterComments);
      });
    }
  }
});

describe("model mistakes", () => {
  for (const [mistake, build] of Object.entries(MISTAKES)) {
    for (const mode of MODES) {
      test(`${mistake} / ${mode}: refused with a reason, or the document still saves`, async () => {
        const fixture = FIXTURE_NAMES[mistake.length % FIXTURE_NAMES.length] ?? "plain";
        const reviewer = await openReviewer(await FIXTURES[fixture]());
        const persisted = persistedFor(reviewer, mode);
        const bridge = createReviewerBridge(reviewer, { mode });
        const blocks = reviewer.getContent() as Block[];
        // Ids the model read from an earlier version of this document.
        const staleIds = (await openReviewer(await FIXTURES.lists()))
          .getContent()
          .map(({ id }) => id);
        const result = call(
          bridge,
          "suggest_changes",
          build(blocks, createRandom(mistake.length), staleIds),
        );
        if (result.ok) {
          const outcome = result.result as SuggestResult;
          if (outcome.applied.length === 0) {
            assert.ok(
              outcome.issues.length > 0,
              `${mistake}: nothing applied and no issue says why`,
            );
          }
        } else {
          assert.ok(result.error.length > 0, `${mistake}: refused without a message`);
        }
        await assertHealthy(reviewer, `${fixture} / ${mode} after ${mistake}`, persisted);
      });
    }
  }
});
