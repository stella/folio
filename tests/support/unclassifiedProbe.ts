// Scratch instrumentation: collect the editor's unclassified-transaction
// reports from each test's page.
import { appendFileSync, mkdirSync } from "node:fs";
import path from "node:path";

import { test as base } from "../support/unclassifiedProbe";

export * from "@playwright/test";

export const UNCLASSIFIED_DIR = path.resolve("test-results/unclassified");

export const test = base.extend<{ unclassifiedProbe: void }>({
  unclassifiedProbe: [
    async ({ page }, use, testInfo) => {
      const records: unknown[] = [];
      const listen = (target: typeof page) =>
        target.on("console", (message) => {
          const text = message.text();
          if (text.startsWith("[folio-unclassified]"))
            records.push(JSON.parse(text.slice("[folio-unclassified]".length)));
        });
      listen(page);
      page.context().on("page", listen);
      await use();
      if (records.length === 0) return;
      mkdirSync(UNCLASSIFIED_DIR, { recursive: true });
      appendFileSync(
        path.join(UNCLASSIFIED_DIR, `${process.pid}.jsonl`),
        records
          .map((record) =>
            JSON.stringify({
              file: path.basename(testInfo.file),
              test: testInfo.titlePath.slice(1).join(" > "),
              record,
            }),
          )
          .join("\n") + "\n",
      );
    },
    { auto: true },
  ],
});
