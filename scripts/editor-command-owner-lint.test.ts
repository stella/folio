import { expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

setDefaultTimeout(30_000);
const ROOT = path.resolve(import.meta.dir, "..");
const MARKER = "folio-editor-commands(command-owner-boundary)";
const lintFixture = (fixture: string) => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.fixtures.config.ts",
      "--no-ignore",
      path.join("test/__fixtures__/editor-commands", fixture),
    ],
    { cwd: ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  expect(output).not.toContain("Failed to load");
  expect(output).not.toContain("Failed to parse");
  return output.split(MARKER).length - 1;
};

test("editor command owner lint covers adapter dispatch and owner registration", () => {
  expect(lintFixture("invalid/packages/react/src/components/DocxEditor.tsx")).toBe(4);
  expect(lintFixture("invalid/packages/vue/src/composables/useFormattingActions.ts")).toBe(1);
  expect(lintFixture("invalid/packages/vue/src/components/Toolbar.vue")).toBe(1);
  expect(lintFixture("valid/packages/react/src/components/DocxEditor.tsx")).toBe(0);
  expect(lintFixture("valid/packages/core/src/controller/hiddenEditorManager.ts")).toBe(0);
});
