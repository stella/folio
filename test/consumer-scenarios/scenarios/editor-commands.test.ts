/**
 * The integrator's editor-command recipes (#1091, #1092), run headlessly
 * through the published prosemirror entry: the command's result saves and
 * reopens with the list the command meant.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import { toggleBulletList, toggleNumberedList } from "@stll/folio-core/prosemirror";

import { packDocument, styleNumberedDocument } from "../support/documents.ts";
import { labelsOf, toggleAndSave } from "../support/editor.ts";

describe("editor list commands in a package without a list of that kind (#1091)", () => {
  test("Numbered List in a package without a numbering part saves", async () => {
    const bytes = await packDocument(fromMarkdown("Intro paragraph.\n\nFirst item\n\nTail."));
    const labels = await labelsOf(await toggleAndSave(bytes, ["First item"], toggleNumberedList));
    assert.deepEqual(labels, ["· Intro paragraph.", "1. First item", "· Tail."]);
  });

  test("Bullet List in a package without a numbering part saves", async () => {
    const bytes = await packDocument(fromMarkdown("Intro.\n\nMake me a bullet"));
    const labels = await labelsOf(
      await toggleAndSave(bytes, ["Make me a bullet"], toggleBulletList),
    );
    assert.deepEqual(labels, ["· Intro.", "• Make me a bullet"]);
  });

  test("Bullet List in a package whose only list is numbered makes a bullet", async () => {
    const bytes = await packDocument(
      fromMarkdown("1. Alpha\n2. Beta\n\nPlain text.\n\nMake me a bullet"),
    );
    const labels = await labelsOf(
      await toggleAndSave(bytes, ["Make me a bullet"], toggleBulletList),
    );
    assert.equal(labels.at(-1), "• Make me a bullet", "the toggled paragraph is not a bullet");
  });
});

describe("a list command after an unrelated list (#1092)", () => {
  test("two paragraphs toggled after prose start a new list at 1", async () => {
    const bytes = await packDocument(
      fromMarkdown(
        "1. Alpha\n2. Beta\n\nUnrelated prose between the lists.\n\nNew list one\n\nNew list two",
      ),
    );
    const labels = await labelsOf(
      await toggleAndSave(bytes, ["New list one", "New list two"], toggleNumberedList),
    );
    assert.deepEqual(
      labels.slice(-2),
      ["1. New list one", "2. New list two"],
      "the toggled paragraphs did not start a new list",
    );
  });
});

describe("a list command in a document whose headings are numbered by style (#1092)", () => {
  test("a body paragraph made a list item leaves the heading numbers alone", async () => {
    const before = await labelsOf(await styleNumberedDocument());
    const after = await labelsOf(
      await toggleAndSave(
        await styleNumberedDocument(),
        ["The Buyer pays on delivery."],
        toggleNumberedList,
      ),
    );
    const headings = (labels: string[]) =>
      labels.filter((label) => /^\d+\.(?:\d+\.)? (?:Scope|Definitions|Payment)$/u.test(label));
    assert.deepEqual(headings(after), headings(before));
    assert.equal(
      after.find((label) => label.endsWith("The Buyer pays on delivery.")),
      "1. The Buyer pays on delivery.",
    );
  });
});
