import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, paragraph, run } from "@stll/folio-core/server";

import { openReviewer, packDocument } from "../support/documents.ts";
import { assertStoryInventory, saveAndReopen } from "../support/invariants.ts";
import { MODES } from "../support/operations.ts";
import { createRandom, sentence, type Random } from "../support/random.ts";

test("note inventory oracle rejects stale, missing and duplicate trailer entries", () => {
  const story = { handle: { type: "footnote", noteId: 37 }, text: "Note." } as const;
  for (const trailer of ["", "[endnote #37] Note.", "[footnote #37] Note.\n[footnote #37] Note."]) {
    assert.throws(
      () =>
        assertStoryInventory({
          listStories: () => [story],
          getNotesAsText: () => trailer,
        }),
      /story trailer and live story inventory disagree/u,
    );
  }
  assert.throws(
    () =>
      assertStoryInventory({
        listStories: () => [],
        getNotesAsText: () => "[footnote #37] Note.",
      }),
    /story trailer and live story inventory disagree/u,
  );
});

test("story inventory oracle treats labels inside story text literally", () => {
  const text = "Quoted marker:\n[footnote #73] A different note.\nSecond paragraph.";
  assertStoryInventory({
    listStories: () => [
      { handle: { type: "header", relationshipId: "rId1" }, text: "Header." },
      { handle: { type: "footnote", noteId: 37 }, text },
    ],
    getNotesAsText: () => `[header first (sections 1, 3)] Header.\n[footnote #37] ${text}`,
  });
});

/** Vary package ids, text, reference order and whether a reference is the last one. */
const noteInventoryDocument = async (random: Random, references: number) => {
  const document = fromMarkdown(sentence(random));
  const id = 1 + random.int(100_000);
  document.package.footnotes = [{ type: "footnote", id, content: [paragraph(sentence(random))] }];
  document.package.endnotes = [{ type: "endnote", id, content: [paragraph(sentence(random))] }];
  const kinds = random.chance(0.5)
    ? (["footnote", "endnote"] as const)
    : (["endnote", "footnote"] as const);
  for (const kind of kinds) {
    for (let index = 0; index < references; index += 1) {
      document.package.document.content.push(
        paragraph([
          run(sentence(random)),
          {
            type: "run",
            content: [{ type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id }],
          },
        ]),
      );
    }
  }
  return { bytes: await packDocument(document), id, kinds };
};

// Enumerate both note kinds, reference cardinalities and every edit mode on
// every run. Check the inventory between individual revision resolutions.
for (const mode of MODES) {
  const resolutions = mode === "direct" ? (["accept"] as const) : (["accept", "reject"] as const);
  for (const resolution of resolutions) {
    for (const references of [1, 2]) {
      test(`note inventories follow ${mode} reference deletions / ${resolution} / ${references} references`, async () => {
        const random = createRandom(mode.length * 97 + resolution.length + references);
        const { bytes, id, kinds } = await noteInventoryDocument(random, references);
        let reviewer = await openReviewer(bytes);
        const original = reviewer.getNotesAsText();
        for (const kind of kinds) {
          // Capture reference paragraphs before rejecting can restore them.
          const blocks = reviewer
            .getContent()
            .filter(
              (candidate) =>
                candidate.kind !== "diagnostic" &&
                candidate.structuralBoundaries?.some(
                  (boundary) => boundary.type === "noteReference" && boundary.noteType === kind,
                ),
            );
          assert.equal(blocks.length, references);
          for (const [index, block] of blocks.entries()) {
            const result = reviewer.applyDocumentOperations(
              {
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                mode,
                operations: [
                  { id: `delete-${kind}-${index}`, type: "deleteBlock", blockId: block.id },
                ],
              },
              { undefinedReferences: "refuse" },
            );
            assert.equal(result.applied.length, 1);
            assertStoryInventory(reviewer);
            if (mode !== "direct") {
              // Resolve reference runs and paragraph marks separately.
              while (reviewer.getChanges().length > 0) {
                const change = reviewer.getChanges().at(0);
                assert.ok(change);
                assert.equal(
                  resolution === "accept"
                    ? reviewer.acceptChange(change)
                    : reviewer.rejectChange(change),
                  true,
                );
                assertStoryInventory(reviewer);
              }
            }
            const kept = (mode !== "direct" && resolution === "reject") || index + 1 < references;
            assert.equal(reviewer.getNotesAsText().includes(`[${kind} #${id}]`), kept);
            reviewer = (
              await saveAndReopen(reviewer, `${mode} / ${resolution} / ${kind} / ${index}`)
            ).reopened;
            assertStoryInventory(reviewer);
          }
        }
        if (mode !== "direct" && resolution === "reject") {
          assert.equal(reviewer.getNotesAsText(), original);
        } else {
          assert.equal(reviewer.getNotesAsText(), "");
        }
      });
    }
  }
}
