import { expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";

setDefaultTimeout(scanTestTimeout(10_000));

const root = path.resolve(import.meta.dir, "../packages/core/src");

test("AI new-list allocation has one resolver and one batch owner", async () => {
  const resolverFiles: string[] = [];
  const newListMintFiles: string[] = [];
  const deferredValueFiles: string[] = [];
  const negativePlaceholderFiles: string[] = [];
  for await (const relative of new Bun.Glob("**/*.ts").scan({ cwd: root })) {
    if (relative.endsWith(".test.ts") || relative.endsWith(".property.test.ts")) continue;
    const source = await Bun.file(path.join(root, relative)).text();
    if (source.includes("resolveNewListOperations")) resolverFiles.push(relative);
    if (source.includes('"newList"') && source.includes("mintListInstance")) {
      newListMintFiles.push(relative);
    }
    if (source.includes('"deferredNewList"')) deferredValueFiles.push(relative);
    if (/placeholder\s*=\s*-\d|numId\s*[<>]=?\s*0/.test(source)) {
      negativePlaceholderFiles.push(relative);
    }
  }

  expect(resolverFiles.toSorted()).toEqual(["ai-edits/apply.ts", "ai-edits/newListNumbering.ts"]);
  expect(newListMintFiles).toEqual(["ai-edits/newListNumbering.ts"]);
  expect(deferredValueFiles).toEqual(["ai-edits/newListNumbering.ts"]);
  expect(negativePlaceholderFiles).toEqual([]);
});

test("deferred new lists are allocated only at will-apply boundaries", async () => {
  const source = await Bun.file(path.join(root, "ai-edits/apply.ts")).text();
  const loop = source.slice(
    source.indexOf("for (const insertionRun of executionGroups)"),
    source.indexOf("const resolvedFinalParagraphDeletions"),
  );
  const beforeSingletonSwitch = loop.slice(0, loop.indexOf("switch (item.operation.type)"));
  expect(beforeSingletonSwitch.indexOf('reason: "pendingParagraphPropertyChange"')).toBeLessThan(
    beforeSingletonSwitch.indexOf('status: "willApply"'),
  );
  expect(beforeSingletonSwitch).toContain('reason: "unsupportedMode"');

  const allocationAfter = (caseName: string, refusal: string, nextCase: string): void => {
    const branch = loop.slice(
      loop.indexOf(`case "${caseName}"`),
      loop.indexOf(`case "${nextCase}"`),
    );
    expect(branch).toContain(refusal);
    expect(branch.indexOf(refusal)).toBeLessThan(branch.indexOf('status: "willApply"'));
  };
  allocationAfter("insertAfterBlock", 'reason: "pendingDeletion"', "insertSignatureTable");
  allocationAfter(
    "setBlockParagraphProperties",
    "pendingParagraphMarkPropertyRefusal",
    "splitBlock",
  );
  allocationAfter("splitBlock", 'reason: "unsupportedBlock"', "mergeBlockWithNext");
  allocationAfter("mergeBlockWithNext", 'reason: "unsupportedBlock"', "commentOnBlock");

  const refusalReasons = [
    "pendingParagraphMarkDeletion",
    "pendingParagraphPropertyChange",
    "pendingRunPropertyChange",
    "protectedReference",
    "unsupportedMode",
    "noopOperation",
  ] as const;
  for (const reason of refusalReasons) expect(source).toContain(reason);
});
