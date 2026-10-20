import { expect, test } from "bun:test";

const REQUIRED_DERIVED_READERS = {
  "packages/core/src/ai-edits/headless.ts": "coveredCommentIdsInProseDoc",
  "packages/core/src/controller/canonicalSession.ts": "survivingComments",
  "packages/core/src/controller/layoutPipeline.ts": "survivingComments",
  "packages/core/src/docx/rezip.ts": "withoutExhaustedComments",
  "packages/core/src/docx/selectiveSave.ts": "withoutExhaustedComments",
  "packages/core/src/export-pdf.ts": "survivingComments",
  "packages/core/src/markdown/index.ts": "withoutExhaustedComments",
} as const;

test("comment readers use the shared exhaustion projection", async () => {
  for (const [path, requiredCall] of Object.entries(REQUIRED_DERIVED_READERS)) {
    expect(await Bun.file(path).text(), path).toContain(requiredCall);
  }
});
