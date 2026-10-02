import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import { createDocx, FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/server";

import { openReviewer } from "../support/documents.ts";
import { projectContentPair, projectSnapshotIdentities } from "../support/identity.ts";
import { saveAndReopen, visibleState } from "../support/invariants.ts";
import { ENABLED_RELATIONS, relationCheckCount, startRelations } from "../support/metamorphic.ts";

type Row = { id: string; idStability?: "stable" | "positional"; text: string; kind: string };

type PairRows = { left: readonly Row[]; right: readonly Row[] };
const assertPair = ({ left, right }: PairRows) => {
  const projected = projectContentPair({ leftRows: left, rightRows: right });
  assert.deepEqual(projected.left, projected.right);
};
const assertPairUnequal = ({ left, right }: PairRows) => {
  const projected = projectContentPair({ leftRows: left, rightRows: right });
  assert.notDeepEqual(projected.left, projected.right);
};
const rowAt = (rows: readonly Row[], index: number): Row => {
  const row = rows.at(index);
  assert.ok(row);
  return row;
};

describe("save comparison block identity", () => {
  test("ordinal and provenance combinations normalize only positional rows", () => {
    for (let mask = 0; mask < 1 << 4; mask += 1) {
      const left = Array.from({ length: 4 }, (_, index): Row => {
        const provenance: Pick<Row, "idStability"> =
          mask & (1 << index) ? { idStability: "positional" } : {};
        return {
          id: `live-${String(index)}`,
          ...provenance,
          kind: "paragraph",
          text: `Text ${String(index)}`,
        };
      });
      const right = left.map((row, index) => {
        const reminted = {
          ...row,
          id: row.idStability === "positional" ? `recomputed-${String(index)}` : row.id,
        };
        if (row.idStability === "positional" && index % 2 === 0) {
          reminted.idStability = "stable";
        } else {
          delete reminted.idStability;
        }
        return reminted;
      });
      assertPair({ left: left, right: right });
      for (const [index, row] of left.entries()) {
        const rightRow = rowAt(right, index);
        const changedId = right.with(index, { ...rightRow, id: `mutated-${String(index)}` });
        if (row.idStability !== "positional") assertPairUnequal({ left: left, right: changedId });
        assertPairUnequal({
          left: left,
          right: right.with(index, { ...rightRow, text: `Changed ${String(index)}` }),
        });
      }
    }
  });

  test("a recomputed positional id cannot hide authored identity, order, count, or content drift", () => {
    const live: Row[] = [
      { id: "authored-a", kind: "paragraph", text: "First" },
      { id: "positional-b", idStability: "positional", kind: "paragraph", text: "Second" },
      { id: "authored-c", kind: "paragraph", text: "Third" },
    ];
    const reminted: Row[] = [
      { id: "authored-a", kind: "paragraph", text: "First" },
      { id: "reminted-b", kind: "paragraph", text: "Second" },
      { id: "authored-c", kind: "paragraph", text: "Third" },
    ];
    assertPair({ left: live, right: reminted });

    assertPairUnequal({
      left: live,
      right: reminted.with(0, { ...rowAt(reminted, 0), id: "changed-authored-a" }),
    });
    assertPairUnequal({ left: live, right: reminted.toSpliced(0, 1) });
    assertPairUnequal({
      left: live,
      right: reminted.toSpliced(0, 2, rowAt(reminted, 1), rowAt(reminted, 0)),
    });
    assertPairUnequal({
      left: live,
      right: reminted.with(1, { ...rowAt(reminted, 1), text: "Changed" }),
    });
  });

  test("resolved text can carry a newly computed positional id", () => {
    const resolved: Row[] = [
      {
        id: "before-resolution",
        idStability: "positional",
        kind: "paragraph",
        text: "Resolved text",
      },
    ];
    const reopened: Row[] = [{ id: "after-resolution", kind: "paragraph", text: "Resolved text" }];
    assertPair({ left: resolved, right: reopened });
    assert.notDeepEqual(
      projectContentPair({ leftRows: resolved, rightRows: reopened }).left,
      reopened,
    );
  });

  test("a repacked positional id may be explicitly marked stable", () => {
    const resolved: Row[] = [
      {
        id: "before-repack",
        idStability: "positional",
        kind: "paragraph",
        text: "Resolved text",
      },
    ];
    const repacked: Row[] = [
      {
        id: "after-repack",
        idStability: "stable",
        kind: "paragraph",
        text: "Resolved text",
      },
    ];
    assertPair({ left: resolved, right: repacked });
  });

  test("a live authored id reused at a positional slot remains visible", () => {
    const live: Row[] = [
      { id: "position", idStability: "positional", kind: "paragraph", text: "Generated" },
    ];
    const reopened: Row[] = [{ id: "authored", kind: "paragraph", text: "Generated" }];
    const projected = projectContentPair({
      leftRows: live,
      rightRows: reopened,
      stableIds: new Set(["authored"]),
    });
    assert.notDeepEqual(projected.left, projected.right);
  });

  test("side maps handle overlaps, permutations, duplicates and authored ids", () => {
    const positional: Row[] = [
      { id: "id-a", idStability: "positional", kind: "paragraph", text: "A" },
      { id: "id-b", idStability: "positional", kind: "paragraph", text: "B" },
    ];
    const overlapPermutation: Row[] = [
      { id: "id-b", kind: "paragraph", text: "A" },
      { id: "id-a", kind: "paragraph", text: "B" },
    ];
    assertPair({ left: positional, right: overlapPermutation });

    const duplicateRight: Row[] = [
      { id: "duplicate", kind: "paragraph", text: "A" },
      { id: "duplicate", kind: "paragraph", text: "B" },
    ];
    assertPairUnequal({ left: positional, right: duplicateRight });

    const authored: Row[] = [
      { id: "authored", kind: "paragraph", text: "Authored" },
      { id: "position", idStability: "positional", kind: "paragraph", text: "Generated" },
    ];
    const rightUsesAuthored: Row[] = [
      { id: "authored", kind: "paragraph", text: "Authored" },
      { id: "authored", kind: "paragraph", text: "Generated" },
    ];
    assertPairUnequal({ left: authored, right: rightUsesAuthored });
  });

  test("bridge snapshot projection maps typed block and anchor identities only", async () => {
    const reviewer = await openReviewer(
      new Uint8Array(await createDocx(fromMarkdown("First paragraph.\n\nSecond paragraph."))),
    );
    const snapshot = reviewer.snapshot();
    const first = snapshot.blocks.at(0);
    const second = snapshot.blocks.at(1);
    assert.ok(first);
    assert.ok(second);
    const anchor = snapshot.anchors[first.id];
    const secondAnchor = snapshot.anchors[second.id];
    assert.ok(anchor);
    assert.ok(secondAnchor);
    const mintedBlock: (typeof snapshot.blocks)[number] = {
      ...first,
      id: "minted",
      idStability: "positional",
      text: "The minted id is prose.",
    };
    const authoredBlock: (typeof snapshot.blocks)[number] = {
      ...second,
      id: "authored",
      idStability: "stable",
    };
    const fixture = {
      ...snapshot,
      blocks: [mintedBlock, authoredBlock],
      anchors: {
        minted: { ...anchor, id: "minted", textHash: "minted" },
        authored: { ...secondAnchor, id: "authored" },
      },
      opaque: "minted",
    };
    const projected = projectSnapshotIdentities(fixture, new Map([["minted", "position:0"]]));
    const projectedBlock = projected.blocks.at(0);
    const stableBlock = projected.blocks.at(1);
    assert.ok(projectedBlock);
    assert.ok(stableBlock);
    assert.equal(projectedBlock.id, "position:0");
    assert.equal(projectedBlock.text, "The minted id is prose.");
    assert.equal(Object.hasOwn(projectedBlock, "idStability"), false);
    assert.equal(stableBlock.id, "authored");
    assert.equal(stableBlock.idStability, "stable");
    assert.equal(projected.anchors["position:0"]?.id, "position:0");
    assert.equal(projected.anchors["position:0"]?.textHash, "minted");
    assert.equal(Reflect.get(projected, "opaque"), "minted");
  });

  test("createDocx without authored paraIds survives tracked text resolution and save", async () => {
    const bytes = new Uint8Array(
      await createDocx(fromMarkdown("First source text.\n\nSecond paragraph.")),
    );
    const reviewer = await openReviewer(bytes);
    const [first] = reviewer.getContent();
    assert.ok(first);
    assert.equal(first.idStability, "positional");
    const receipt = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "resolve-positional-text",
          type: "replaceInBlock",
          blockId: first.id,
          find: "source",
          replace: "resolved",
        },
      ],
    });
    assert.equal(receipt.applied.length, 1);
    reviewer.acceptAll();
    const resolved = reviewer.getContent().find(({ text }) => text === "First resolved text.");
    assert.ok(resolved);
    assert.equal(resolved.idStability, "positional");
    const expected = visibleState(reviewer);
    const result = await saveAndReopen(reviewer, "resolved positional text", {
      persisted: expected,
    });
    assert.deepEqual(
      visibleState(result.reopened).blocks.map(({ text }) => text),
      ["First resolved text.", "Second paragraph."],
    );
  });

  test(
    "readerStability afterStep compares resolved tracked content through shared identity maps",
    { skip: !ENABLED_RELATIONS.has("readerStability") },
    async () => {
      const fixture = new Uint8Array(
        await createDocx(fromMarkdown("First source text.\n\nSecond paragraph.")),
      );
      const reviewer = await openReviewer(fixture);
      const first = reviewer.getContent().at(0);
      assert.ok(first);
      const before = relationCheckCount("readerStability");
      const relations = await startRelations({
        fixture,
        reviewer,
        mode: "tracked-changes",
        seed: 71,
      });
      const receipt = reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "reader-stability-resolution",
            type: "replaceInBlock",
            blockId: first.id,
            find: "source",
            replace: "resolved",
          },
        ],
      });
      assert.equal(receipt.applied.length, 1);
      reviewer.acceptAll();
      const saved = await saveAndReopen(reviewer, "reader stability positional ids", {
        compare: false,
      });
      await relations.afterStep(reviewer, saved, "reader stability positional ids");
      assert.equal(relationCheckCount("readerStability"), before + 1);
    },
  );
});
