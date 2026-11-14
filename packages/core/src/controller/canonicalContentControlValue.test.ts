import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "../canonical/canonicalSession";
import {
  setContentControlValue,
  type SetContentControlValueInput,
} from "../content-controls";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import { CANONICAL_RULE } from "../types/canonicalCapabilities";
import type { BlockContent, Document, SdtProperties } from "../types/document";
import { prepareCanonicalCommands } from "./canonicalStructure";

setDefaultTimeout(propertyTestTimeout(30_000));

type Lock = "contentLocked" | "sdtContentLocked" | "sdtLocked";

const LIST = [
  { displayText: "California", value: "ca" },
  { displayText: "New York", value: "ny" },
];

const CONTROLS = {
  dropdown: { sdtType: "dropdown", listItems: LIST, dropdownLastValue: "ca" },
  comboBox: { sdtType: "comboBox", listItems: LIST, dropdownLastValue: "ca" },
  date: { sdtType: "date", dateFormat: "yyyy-MM-dd" },
  checkbox: { sdtType: "checkbox", checked: false },
} as const satisfies Record<string, SdtProperties>;
type Kind = keyof typeof CONTROLS;

const SHOWN: Record<Kind, string> = {
  dropdown: "California",
  comboBox: "California",
  date: "2026-01-15",
  checkbox: "☐",
};

/** One typed control between two plain paragraphs. */
const controlDocument = (kind: Kind, lock?: Lock): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "10000000",
          content: [
            { type: "run", content: [{ type: "text", text: "Before" }] },
          ],
        },
        {
          type: "blockSdt",
          properties: {
            ...structuredClone(CONTROLS[kind]),
            id: 7,
            tag: kind,
            ...(lock === undefined ? {} : { lock }),
          },
          content: [
            {
              type: "paragraph",
              paraId: "20000000",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: SHOWN[kind] }],
                },
              ],
            },
          ],
        },
        {
          type: "paragraph",
          paraId: "30000000",
          content: [
            { type: "run", content: [{ type: "text", text: "After" }] },
          ],
        },
      ],
    },
  },
});

const INPUTS: Record<Kind, SetContentControlValueInput> = {
  dropdown: { kind: "dropdown", value: "ny" },
  comboBox: { kind: "dropdown", value: "ny" },
  date: { kind: "date", date: "2026-06-02" },
  checkbox: { kind: "checkbox", checked: true },
};

const session = async (kind: Kind, lock?: Lock): Promise<CanonicalSession> =>
  createCanonicalSession(
    await parseDocx(
      (
        await prepareCanonicalDocxInput(
          await createDocx(controlDocument(kind, lock)),
        )
      ).unwrap(),
      { preloadFonts: false, detectVariables: false },
    ),
  ).unwrap();

const stateFor = (current: CanonicalSession) =>
  EditorState.create({ schema, doc: current.projection.doc });
const accept = (state: EditorState, commit: CanonicalCommit) => {
  const next = state.apply(commit.transaction);
  expect(commit.publish().isOk()).toBe(true);
  return next;
};
const controlAt = (state: EditorState): number => {
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found === undefined && node.type.name === "blockSdt") found = position;
    return found === undefined;
  });
  return found ?? panic("No control is projected.");
};
const control = (document: Document) => {
  const block = document.package.document.content.find(
    (item): item is Extract<BlockContent, { type: "blockSdt" }> =>
      item.type === "blockSdt",
  );
  return block ?? panic("The control is gone.");
};
const runsOf = (item: Paragraph["content"][number]): Run[] => {
  if (item.type === "run") return [item];
  if (item.type === "insertion")
    return item.content.filter((child): child is Run => child.type === "run");
  return [];
};
/** The control's shown text and symbols, tracked deletions left out. */
const shownText = (document: Document): string =>
  control(document)
    .content.flatMap((block) =>
      block.type === "paragraph"
        ? block.content.flatMap((item) => {
            return runsOf(item).flatMap((run) =>
              run.content.flatMap((leaf) => {
                if (leaf.type === "text") return [leaf.text];
                // A checkbox shows its state as a symbol.
                return leaf.type === "symbol" ? [JSON.stringify(leaf)] : [];
              }),
            );
          })
        : [],
    )
    .join("");
/** The properties a widget sets, as a package states them. */
const value = (properties: SdtProperties) => ({
  dropdownLastValue: properties.dropdownLastValue,
  checked: properties.checked,
  dateValueISO: properties.dateValueISO?.slice(0, 10),
});

const pick = (current: CanonicalSession, state: EditorState, kind: Kind) =>
  prepareCanonicalCommands(current, state, [
    {
      type: "setContentControlValue",
      at: controlAt(state),
      input: INPUTS[kind],
    },
  ]);

const save = async (current: CanonicalSession) =>
  parseDocx(
    (
      await serializeCanonicalSave({
        snapshot: current.captureSaveSnapshot(),
        featureFlags: { selectiveSave: true },
      })
    ).buffer,
    { preloadFonts: false },
  );

describe("a typed control's value in the canonical session", () => {
  for (const kind of Object.keys(CONTROLS) as Kind[]) {
    test(`${kind}: one undo step sets and restores the value and its text, and save keeps it`, async () => {
      const current = await session(kind);
      const before = current.document;
      let state = accept(
        stateFor(current),
        pick(current, stateFor(current), kind).unwrap(),
      );
      const set = control(current.document);
      expect(shownText(current.document)).not.toBe(SHOWN[kind]);
      const reopened = await save(current);
      expect(value(control(reopened).properties)).toEqual(
        value(set.properties),
      );
      expect(shownText(reopened)).toBe(shownText(current.document));
      state = accept(state, current.prepareUndo(state).unwrap());
      expect(current.document).toEqual(before);
      expect(current.canUndo).toBe(false);
      accept(state, current.prepareRedo(state).unwrap());
      expect(control(current.document).properties).toEqual(set.properties);
    });
  }

  test("suggesting mode tracks the shown text and sets the value", async () => {
    const current = await session("dropdown");
    current.setMode({ type: "suggesting", author: "Reviewer" });
    const state = stateFor(current);
    accept(state, pick(current, state, "dropdown").unwrap());
    const block = control(current.document);
    expect(block.properties.dropdownLastValue).toBe("ny");
    const paragraph = block.content.at(0);
    const kinds =
      paragraph?.type === "paragraph"
        ? paragraph.content.map((item) => item.type)
        : [];
    expect(kinds).toContain("insertion");
    expect(kinds).toContain("deletion");
  });

  for (const lock of ["contentLocked", "sdtContentLocked"] as const) {
    test(`a ${lock} control refuses with the lock rule and changes nothing`, async () => {
      const current = await session("dropdown", lock);
      const before = current.document;
      const refused = pick(current, stateFor(current), "dropdown");
      if (refused.isOk()) return panic("A locked control took a value.");
      expect(refused.error.ground).toEqual({
        rule: CANONICAL_RULE.contentControlLock,
      });
      expect(current.document).toBe(before);
    });
  }

  test("an sdtLocked control still takes a value", async () => {
    const current = await session("dropdown", "sdtLocked");
    const state = stateFor(current);
    accept(state, pick(current, state, "dropdown").unwrap());
    expect(control(current.document).properties.dropdownLastValue).toBe("ny");
  });

  test(
    "the saved package states what the headless rules state, for every kind",
    async () => {
      await assertProperty(
        fc.asyncProperty(
          fc.constantFrom(...(Object.keys(CONTROLS) as Kind[])),
          fc.constantFrom("ca", "ny"),
          fc.date({
            min: new Date("2000-01-01T00:00:00Z"),
            max: new Date("2099-12-31T00:00:00Z"),
            noInvalidDate: true,
          }),
          fc.boolean(),
          async (kind, choice, day, checked) => {
            let input: SetContentControlValueInput = {
              kind: "dropdown",
              value: choice,
            };
            if (kind === "date")
              input = { kind: "date", date: day.toISOString().slice(0, 10) };
            if (kind === "checkbox") input = { kind: "checkbox", checked };
            const current = await session(kind);
            const state = stateFor(current);
            const prepared = prepareCanonicalCommands(current, state, [
              { type: "setContentControlValue", at: controlAt(state), input },
            ]);
            if (prepared.isOk()) accept(state, prepared.value);
            else expect(prepared.error.reason).toBe("noChange");
            const canonical = await save(current);
            const headless = await parseDocx(
              await createDocx(
                setContentControlValue(controlDocument(kind), { id: 7 }, input),
              ),
              { preloadFonts: false },
            );
            expect(value(control(canonical).properties)).toEqual(
              value(control(headless).properties),
            );
            expect(shownText(canonical)).toBe(shownText(headless));
          },
        ),
        {
          id: "canonical content control values save as the headless rules state",
          numRuns: 12,
        },
      );
    },
    propertyTestTimeout(60_000),
  );
});
