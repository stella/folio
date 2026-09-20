import { describe, expect, test } from "bun:test";

import { ExtensionManager } from "./ExtensionManager";
import { singletonManager } from "../schema";

// A looked-up command keeps its own signature: proved at the type level in
// `typecheck/prosemirror/commandRegistry.typecheck.ts`.
describe("Folio command registry", () => {
  test("requires every built-in wrapper command at singleton startup", () => {
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
