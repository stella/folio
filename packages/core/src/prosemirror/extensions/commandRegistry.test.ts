import { describe, expect, test } from "bun:test";

import { ExtensionManager } from "./ExtensionManager";
import { singletonManager } from "../schema";

// Compile-time contract checks: this function is intentionally never called.
// The rejected calls must remain errors if lookup typing is weakened.
const assertStaticCommandTypes = (): void => {
  const typedCommand = singletonManager.getCommand("setLineSpacing");
  if (typedCommand) {
    typedCommand(240, "auto");
    // @ts-expect-error line spacing requires a number first
    typedCommand("240");
  }

  const typedTableCommand = singletonManager.requireCommand("setTableBorderPreset");
  // @ts-expect-error table border presets are not arbitrary strings
  typedTableCommand("not-a-preset");
};

describe("Folio command registry", () => {
  test("requires every built-in wrapper command at singleton startup", () => {
    expect(assertStaticCommandTypes).toBeFunction();
    expect(singletonManager.requireCommand("setLineSpacing")).toBeFunction();
    expect(singletonManager.requireCommand("setTableBorderPreset")).toBeFunction();
  });

  test("reports missing commands instead of returning an unchecked assertion", () => {
    const manager = new ExtensionManager([]);
    expect(() => manager.requireCommand("extension-only-command")).toThrow(
      'command "extension-only-command" is not registered',
    );
  });
});
