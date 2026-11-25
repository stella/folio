import { expect, test } from "bun:test";

import { createReferenceEditorKit } from "../../__tests__/createReferenceEditorKit";
import { schema } from "../schema";
import { createStarterKit } from "./StarterKit";

test("StarterKit keeps commands, keymaps and clipboard transforms", () => {
  const viewRuntimes = createStarterKit({
    historyDepth: 5,
    historyNewGroupDelay: 10,
  }).map((extension) => extension.onSchemaReady({ schema }));
  const referenceRuntimes = createReferenceEditorKit().map((extension) =>
    extension.onSchemaReady({ schema }),
  );
  const clipboardProps = new Set([
    "transformPastedHTML",
    "transformPasted",
    "transformCopied",
    "clipboardTextSerializer",
  ]);
  const clipboardHooks = (runtimes: typeof viewRuntimes) =>
    runtimes.flatMap(({ plugins }) =>
      (plugins ?? []).flatMap((plugin) =>
        Object.keys(plugin.props).filter((prop) => clipboardProps.has(prop)),
      ),
    );

  for (const runtime of viewRuntimes)
    for (const plugin of runtime.plugins ?? []) {
      expect(Object.keys(plugin.spec)).toEqual(["props"]);
      for (const prop of Object.keys(plugin.props))
        expect(clipboardProps.has(prop)).toBe(true);
    }
  expect(clipboardHooks(viewRuntimes)).toContain("transformPasted");
  expect(clipboardHooks(viewRuntimes)).toContain("transformPastedHTML");
  expect(clipboardHooks(viewRuntimes)).toEqual(
    clipboardHooks(referenceRuntimes),
  );
  expect(
    viewRuntimes.some(({ commands }) => commands?.["undo"] !== undefined),
  ).toBe(true);
  expect(
    viewRuntimes.some(
      ({ keyboardShortcuts }) => keyboardShortcuts?.["Mod-z"] !== undefined,
    ),
  ).toBe(true);
});
