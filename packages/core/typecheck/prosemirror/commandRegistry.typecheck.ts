/**
 * Compile-time proof that a looked-up command keeps its own signature.
 *
 * `getCommand` and `requireCommand` are overloaded on the command name, so an
 * extension command arrives typed rather than as a bare callable: calling one
 * with the wrong arguments must stay a type error. The claim used to sit in
 * `extensions/commandRegistry.test.ts`, inside a function the suite only
 * asserted was a function, where nothing checked it — a package's `typecheck`
 * runs over `tsconfig.build.json`, which excludes `*.test.ts`.
 */

import { singletonManager } from "../../src/prosemirror/schema";

const lineSpacing = singletonManager.getCommand("setLineSpacing");
if (lineSpacing !== undefined) {
  lineSpacing(240, "auto");
  // @ts-expect-error line spacing requires a number first
  lineSpacing("240");
}

const tableBorderPreset = singletonManager.requireCommand("setTableBorderPreset");
// @ts-expect-error table border presets are not arbitrary strings
tableBorderPreset("not-a-preset");

export type CommandLookupProof = [typeof lineSpacing, typeof tableBorderPreset];
