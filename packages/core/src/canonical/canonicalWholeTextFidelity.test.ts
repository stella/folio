import { expect, spyOn, test } from "bun:test";

import * as native from "@stll/docx-core/ops";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import * as fidelity from "./canonicalWholeTextFidelity";

test.each([16, 64])(
  "verifies %i isolated whole-text replacements in two native resolutions",
  async (count) => {
    const template = createEmptyDocument();
    const content: Paragraph[] = Array.from({ length: 1024 }, (_, i) => ({
      type: "paragraph",
      paraId: (i + 1).toString(16).padStart(8, "0"),
      content: [{ type: "run", content: [{ type: "text", text: `P ${i}` }] }],
    }));
    const input = await createDocx({
      ...template,
      package: { ...template.package, document: { content } },
    });
    const open = () =>
      FolioDocxReviewer.fromBuffer(input, {
        session: "canonical",
        author: "Review",
      });
    const reviewer = await open();
    const snapshot = reviewer.snapshot();
    const ops = snapshot.blocks
      .filter((_, index) => index % (1024 / count) === 2)
      .map((block, i) => ({
        type: "replaceInBlock" as const,
        id: String(i),
        blockId: block.id,
        find: block.text,
        replace: `Changed ${i}`,
      }));
    const proof = spyOn(fidelity, "verifyWholeTextBatch");
    const apply = spyOn(native, "applyDocumentOps");
    const run = (session: typeof reviewer) =>
      session.applyOperations(ops, {
        mode: "tracked-changes",
        snapshot,
        revisionStamp: { idSeed: 100, date: "2026-03-01T00:00:00Z" },
      });
    const work = () =>
      apply.mock.calls.flatMap(([document, operations]) =>
        operations
          .filter((op) => op.type === native.DOCUMENT_OP_TYPES.RESOLVE_REVISION)
          .map((op) => ({
            decision: op.decision,
            width: document.package.document.content.length,
            ids: op.revisionIds.length,
          })),
      );
    const inverseByParagraph = () => {
      const byParagraph = new Map<string, native.DocumentOp[]>();
      for (const [index, [, operations]] of apply.mock.calls.entries()) {
        const decision = operations.at(0);
        if (decision?.type !== native.DOCUMENT_OP_TYPES.RESOLVE_REVISION)
          continue;
        const edit = apply.mock.results.at(index)?.value;
        if (!edit?.isOk()) throw new Error("Expected a native resolution");
        for (const op of edit.value.inverse) {
          if (op.type !== native.DOCUMENT_OP_TYPES.REPLACE_INLINE)
            throw new Error("Expected an inline inverse");
          const key = decision.decision + "/" + op.blockId;
          const group = byParagraph.get(key) ?? [];
          group.push(op);
          byParagraph.set(key, group);
        }
      }
      return byParagraph;
    };
    try {
      const result = run(reviewer);
      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id)).toEqual(
        ops.map(({ id }) => id),
      );
      expect(proof.mock.results.map(({ value }) => value)).toEqual([
        { type: "verified" },
      ]);
      expect(work()).toEqual(
        ["reject", "accept"].map((decision) => ({
          decision,
          width: count,
          ids: count * 2,
        })),
      );

      const groupedInverse = inverseByParagraph();
      expect(groupedInverse.size).toBe(count * 2);
      const options = proof.mock.calls.at(0)?.[0];
      const first = options?.changes.at(0);
      if (options === undefined || first === undefined)
        throw new Error("Expected a checked paragraph");
      expect(
        fidelity.verifyWholeTextBatch({ ...options, changes: [first, first] }),
      ).toEqual({ type: "fallback", reason: "sharedIdentity" });
      expect(
        fidelity.verifyWholeTextBatch({
          ...options,
          changes: [
            first,
            {
              ...first,
              before: { ...first.before, formatting: { alignment: "center" } },
            },
          ],
        }),
      ).toEqual({ type: "fallback", reason: "shape" });
      expect(
        fidelity.verifyWholeTextBatch({ ...options, resolve: () => undefined }),
      ).toEqual({ type: "fallback", reason: "resolution" });

      expect(
        fidelity.verifyWholeTextBatch({ ...options, changes: [first] }),
      ).toEqual({ type: "fallback", reason: "singleChange" });
      expect(
        fidelity.verifyWholeTextBatch({
          ...options,
          resolve: () =>
            options.changes.map(({ before }) => {
              if (before.paraId === undefined)
                throw new Error("Missing fixture paragraph id");
              return {
                type: "paragraph",
                paraId: before.paraId,
                content: [
                  {
                    type: "run",
                    content: [{ type: "text", text: "different" }],
                  },
                ],
              };
            }),
        }),
      ).toEqual({ type: "fallback", reason: "content" });
      const reference = await open();
      proof.mockClear();
      proof.mockReturnValue({ type: "fallback", reason: "shape" });
      apply.mockClear();
      const ordinary = run(reference);
      expect(proof).toHaveBeenCalledTimes(1);
      expect(work().length).toBe(count * 2);
      expect(work().every(({ width, ids }) => width === 1 && ids === 2)).toBe(
        true,
      );
      expect(inverseByParagraph()).toEqual(groupedInverse);
      expect(ordinary).toEqual(result);
      expect(reference.snapshot().blocks).toEqual(reviewer.snapshot().blocks);
    } finally {
      apply.mockRestore();
      proof.mockRestore();
    }
  },
);
