import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createEmptyDocument } from "../utils/createDocument";
import { EDITOR_MODES } from "./editorHarness";

const stepArbitrary = fc.oneof(
  fc.record({
    type: fc.constant("type" as const),
    text: fc.constantFrom("a", "é", "😀"),
  }),
  fc.record({
    type: fc.constantFrom(
      "undo" as const,
      "redo" as const,
      "save" as const,
      "sectionStart" as const,
    ),
  }),
);

const sourceDocument = () => {
  const source = createEmptyDocument({ initialText: "pending" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph")
    return panic("The pending-changes fixture has no paragraph.");
  paragraph.paraId = "12345678";
  return source;
};

test(
  "a canonical session reports pending changes exactly when a change was committed since it opened or was last saved",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom(...EDITOR_MODES),
        fc.integer({ min: 1, max: 8 }),
        fc.array(stepArbitrary, { minLength: 1, maxLength: 14 }),
        async (mode, caret, steps) => {
          const driver = createCanonicalEditorHarness(sourceDocument(), mode);
          try {
            const api = driver.history;
            api.setSelection(caret);
            // What the host can observe: an edit it made, or an undo or redo
            // that reported a change, since the last save it made.
            let committed = false;
            expect(api.hasCanonicalPendingChanges()).toBe(false);
            for (const [index, step] of steps.entries()) {
              switch (step.type) {
                case "type":
                  driver.typeText(step.text);
                  committed = true;
                  break;
                case "undo":
                  if (api.undo()) committed = true;
                  break;
                case "redo":
                  if (api.redo()) committed = true;
                  break;
                case "sectionStart": {
                  const applied = api.applyCanonicalSectionProperties({
                    footnotePr: { numStart: index + 2 },
                  });
                  expect(applied?.status).toBe("applied");
                  committed = true;
                  break;
                }
                case "save": {
                  const snapshot =
                    api.captureCanonicalSave() ??
                    panic("A canonical session must capture its save.");
                  expect(api.isCanonicalSaveCurrent(snapshot.version)).toBe(
                    true,
                  );
                  api.markCanonicalSaved(snapshot.version);
                  committed = false;
                  break;
                }
                default: {
                  const exhaustive: never = step;
                  return exhaustive;
                }
              }
              expect(driver.refusals).toEqual([]);
              expect(api.hasCanonicalPendingChanges()).toBe(committed);
            }
          } finally {
            driver.dispose();
          }
        },
      ),
      {
        id: "a canonical session reports pending changes exactly when a change was committed since it opened or was last saved",
        numRuns: 40,
      },
    );
  },
  propertyTestTimeout(60_000),
);
