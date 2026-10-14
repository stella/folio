import { expect, test } from "bun:test";
import {
  DISPLAY_MODES,
  getEditorReadOnlyReason,
  getReviewViewReadOnlyReason,
} from "./EditorModeManager";
import type { DisplayMode, EditorMode, EditorReadOnlyReason } from "./EditorModeManager";

const modeCases = {
  editing: "editing",
  suggesting: "suggesting",
  viewing: "viewing",
} as const satisfies Record<EditorMode, EditorMode>;
const viewReasons = {
  "all-markup": null,
  "simple-markup": "review-display-mode",
  "no-markup": "review-display-mode",
  original: "review-display-mode",
} as const satisfies Record<DisplayMode, EditorReadOnlyReason | null>;

test("every display and editing mode derives an explicit read-only reason", () => {
  for (const displayMode of DISPLAY_MODES) {
    expect(getReviewViewReadOnlyReason(displayMode)).toBe(viewReasons[displayMode]);
    for (const editingMode of Object.values(modeCases)) {
      for (const readOnlyProp of [false, true]) {
        const hostReason = readOnlyProp ? "host-read-only" : null;
        const modeReason = editingMode === "viewing" ? "viewing-mode" : null;
        const expected = viewReasons[displayMode] ?? hostReason ?? modeReason;
        expect(getEditorReadOnlyReason({ displayMode, editingMode, readOnlyProp })).toBe(expected);
      }
    }
  }
});
