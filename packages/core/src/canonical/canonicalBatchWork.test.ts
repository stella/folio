import { expect, spyOn, test } from "bun:test";

import * as nativeOps from "@stll/docx-core/ops";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";

test("tracked insertion batches bound story visits and review resolutions", async () => {
  const template = createEmptyDocument();
  const input = await createDocx({
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content: [] },
    },
  });
  const work = async (count: number) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(input, {
      session: "canonical",
    });
    const snapshot = reviewer.snapshot();
    const anchor = snapshot.blocks.at(0);
    if (anchor === undefined) throw new Error("Expected the blank paragraph");
    const walks = spyOn(nativeOps, "storyParagraphs");
    const apply = spyOn(nativeOps, "applyDocumentOps");
    try {
      const result = reviewer.applyOperations(
        Array.from({ length: count }, (_, i) => ({
          id: String(i),
          type:
            i === 0 ? ("replaceBlock" as const) : ("insertAfterBlock" as const),
          blockId: anchor.id,
          text: String(i),
        })),
        { mode: "tracked-changes", snapshot },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied.length).toBe(count);
      // Every block in this fixture is a paragraph. Count the lengths of the
      // actual lists traversed, including the native Accept All and Reject All.
      const visits = walks.mock.calls.reduce(
        (total, [body]) => total + body.content.length,
        0,
      );
      const resolutions = apply.mock.calls
        .flatMap(([, ops]) => ops)
        .filter(
          (op) => op.type === nativeOps.DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        ).length;
      return { visits, resolutions };
    } finally {
      apply.mockRestore();
      walks.mockRestore();
    }
  };
  const small = await work(16);
  const large = await work(64);
  expect(small.resolutions).toBeGreaterThan(0);
  expect(large.resolutions).toBe(small.resolutions);
  expect(small.visits).toBeGreaterThan(0);
  expect(large.visits).toBeLessThanOrEqual(5 * small.visits);
});
