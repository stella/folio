import { expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";

setDefaultTimeout(scanTestTimeout(10_000));

const root = path.resolve(import.meta.dir, "../packages/core/src");

test("AI new-list allocation has one resolver and one batch owner", async () => {
  const resolverFiles: string[] = [];
  const newListMintFiles: string[] = [];
  for await (const relative of new Bun.Glob("**/*.ts").scan({ cwd: root })) {
    if (relative.endsWith(".test.ts") || relative.endsWith(".property.test.ts")) continue;
    const source = await Bun.file(path.join(root, relative)).text();
    if (source.includes("resolveNewListOperations")) resolverFiles.push(relative);
    if (source.includes('"newList"') && source.includes("mintListInstance")) {
      newListMintFiles.push(relative);
    }
  }

  expect(resolverFiles.toSorted()).toEqual(["ai-edits/apply.ts", "ai-edits/newListNumbering.ts"]);
  expect(newListMintFiles).toEqual(["ai-edits/newListNumbering.ts"]);
});
