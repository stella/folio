/**
 * Budget for an in-process test whose writes sync to disk (`FileHandle.sync`,
 * the CLI's journal, sidecar and output writes). Syncing is the durability a
 * write promises, so tests keep it; its latency on a shared CI runner has no
 * bound, so such a test gets this budget instead of the test runner's 5 s.
 * `scripts/disk-sync-test-timeout.test.ts` binds every test reaching a sync.
 */
export const DISK_SYNC_TEST_TIMEOUT_MS = 30_000;
