import { panic } from "better-result";
import { expect, setSystemTime, test } from "bun:test";
import fc from "fast-check";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { assertExactModel } from "../../../../test/exactModel";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createEmptyDocument } from "../utils/createDocument";
import { EDITOR_MODES } from "./editorHarness";

// Past the journal's grouping delay, so a step marked as one starts a new
// undo group in both editors alike.
const GROUP_GAP_MS = 60_000;

const stepArbitrary = fc.oneof(
  fc.record({
    type: fc.constant("type" as const),
    text: fc.constantFrom("a", "é", "東", "😀", "xy"),
    newGroup: fc.boolean(),
  }),
  fc.record({ type: fc.constantFrom("undo" as const, "redo" as const) }),
);

const sourceDocument = () => {
  const source = createEmptyDocument({ initialText: "history" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph")
    return panic("The history fixture has no paragraph.");
  paragraph.paraId = "12345678";
  return source;
};

test(
  "the registry's undo and redo commands step a canonical session's journal as its own undo and redo do",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom(...EDITOR_MODES),
        fc.integer({ min: 1, max: 8 }),
        fc.array(stepArbitrary, { minLength: 1, maxLength: 16 }),
        async (mode, caret, steps) => {
          let now = Date.UTC(2026, 0, 1);
          setSystemTime(new Date(now));
          // `commands` runs the registered commands; `journal` steps the
          // session's history directly, as the undo keys and the host API do.
          const commands = createCanonicalEditorHarness(sourceDocument(), mode);
          const journal = createCanonicalEditorHarness(sourceDocument(), mode);
          try {
            for (const driver of [commands, journal])
              driver.history.setSelection(caret);
            for (const step of steps) {
              let results: readonly [boolean, boolean] = [true, true];
              switch (step.type) {
                case "type":
                  if (step.newGroup) {
                    now += GROUP_GAP_MS;
                    setSystemTime(new Date(now));
                  }
                  commands.typeText(step.text);
                  journal.typeText(step.text);
                  break;
                case "undo":
                case "redo": {
                  const command =
                    commands.commandManager.getCommand(step.type)?.() ??
                    panic(`The registry has no ${step.type} command.`);
                  results = [
                    commands.execute(command),
                    step.type === "undo"
                      ? journal.history.undo()
                      : journal.history.redo(),
                  ];
                  break;
                }
                default: {
                  const exhaustive: never = step;
                  return exhaustive;
                }
              }
              expect(commands.refusals).toEqual([]);
              expect(journal.refusals).toEqual([]);
              expect(results[0]).toBe(results[1]);
              assertExactModel(commands.snapshot(), journal.snapshot());
              expect(commands.state.selection.toJSON()).toEqual(
                journal.state.selection.toJSON(),
              );
              expect([
                commands.history.canUndo(),
                commands.history.canRedo(),
              ]).toEqual([
                journal.history.canUndo(),
                journal.history.canRedo(),
              ]);
            }
          } finally {
            commands.dispose();
            journal.dispose();
            setSystemTime();
          }
        },
      ),
      {
        id: "the registry's undo and redo commands step a canonical session's journal as its own undo and redo do",
        numRuns: 40,
      },
    );
  },
  propertyTestTimeout(60_000),
);
