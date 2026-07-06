import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

// Both adapters now wire `createAIEditSnapshot` over the shared
// `@stll/folio-core/ai-edits` snapshot builder. Driving it through the bridge
// proves the AI-edit snapshot surface is present and produces a non-empty block
// list in each adapter — the precondition the AI chat composer gates "send" on.
// Because both playgrounds parse the same fixture with the same core builder,
// the block count is deterministic and identical across React and Vue.
forEachAdapter("ai-edit: createAIEditSnapshot yields blocks", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const blockCount = await page.evaluate(() => window.__folioParity?.aiSnapshotBlockCount() ?? -1);

  expect(blockCount).toBeGreaterThan(0);
});
