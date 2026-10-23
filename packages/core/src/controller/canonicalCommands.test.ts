import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { type Command, EditorState, TextSelection } from "prosemirror-state";

import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { marksToTextFormatting } from "../prosemirror/runFormattingFromMarks";
import { schema, singletonManager } from "../prosemirror/schema";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";
import { prepareCanonicalCommands } from "./canonicalStructure";

describe("canonical command integration", () => {
  test("every public mark command matches legacy authored formatting over mixed runs", () => {
    const commands = [
      singletonManager.requireCommand("toggleBold")(),
      singletonManager.requireCommand("toggleItalic")(),
      singletonManager.requireCommand("toggleUnderline")(),
      singletonManager.requireCommand("setUnderlineStyle")("double", {
        rgb: "654321",
      }),
      singletonManager.requireCommand("toggleStrike")(),
      singletonManager.requireCommand("toggleSuperscript")(),
      singletonManager.requireCommand("toggleSubscript")(),
      singletonManager.requireCommand("setFontSize")(28),
      singletonManager.requireCommand("clearFontSize")(),
      singletonManager.requireCommand("setFontFamily")("New Font"),
      singletonManager.requireCommand("clearFontFamily")(),
      singletonManager.requireCommand("setTextColor")({ rgb: "123456" }),
      singletonManager.requireCommand("clearTextColor")(),
      singletonManager.requireCommand("setHighlight")("yellow"),
      singletonManager.requireCommand("clearHighlight")(),
    ];
    for (const enabled of [false, true]) {
      for (const reverse of [false, true]) {
        for (const command of commands) {
          const marks = (suffix: string) =>
            enabled
              ? [
                  schema.mark("bold"),
                  schema.mark("italic"),
                  schema.mark("underline", {
                    style: "single",
                    color: { rgb: suffix },
                  }),
                  schema.mark("strike"),
                  schema.mark("fontSize", { size: 22 }),
                  schema.mark("fontFamily", {
                    ascii: `Old ${suffix}`,
                    eastAsia: `East ${suffix}`,
                  }),
                  schema.mark("textColor", { rgb: suffix }),
                  schema.mark("highlight", { color: "green" }),
                  schema.mark("superscript"),
                ]
              : [];
          const doc = schema.node("doc", null, [
            schema.node("paragraph", { paraId: "12345678" }, [
              schema.text("ab", marks("112233")),
              schema.text("cd", marks("445566")),
            ]),
          ]);
          const session = createCanonicalSession({
            package: { document: { content: proseDocToBlocks(doc) } },
          }).unwrap();
          const state = EditorState.create({
            schema,
            doc: session.projection.doc,
            selection: TextSelection.create(
              session.projection.doc,
              reverse ? 5 : 1,
              reverse ? 1 : 5,
            ),
          });
          let legacy = state;
          expect(
            command(state, (transaction) => {
              legacy = state.apply(transaction);
            }),
          ).toBe(true);
          const intents = getCanonicalCommandIntents(command, state);
          if (intents === undefined)
            throw new TypeError(
              "A public mark command has no canonical descriptor.",
            );
          const prepared = prepareCanonicalCommands(session, state, intents);
          let canonical = state;
          if (prepared.isErr()) expect(prepared.error.reason).toBe("noChange");
          else
            canonical = publishCanonicalProjection({
              session,
              state,
              commit: prepared.value,
            }).unwrap().state;
          const formatting = (value: EditorState) => {
            const characters: {
              text: string;
              formatting: ReturnType<typeof marksToTextFormatting>;
            }[] = [];
            value.doc.descendants((node) => {
              if (!node.isText) return;
              const authored = marksToTextFormatting(node.marks);
              for (const text of node.text ?? "")
                characters.push({ text, formatting: authored });
            });
            return characters;
          };
          expect(formatting(canonical)).toEqual(formatting(legacy));
          expect(canonical.selection.toJSON()).toEqual(
            legacy.selection.toJSON(),
          );
        }
      }
    }
  });

  test("toggles on a partly formatted selection clear all of it, then apply to all of it", () => {
    const toggles = [
      ["bold", "toggleBold"],
      ["italic", "toggleItalic"],
      ["underline", "toggleUnderline"],
      ["strike", "toggleStrike"],
      ["superscript", "toggleSuperscript"],
      ["subscript", "toggleSubscript"],
    ] as const;
    const runLegacy = (state: EditorState, command: Command): EditorState => {
      let next: EditorState | undefined;
      command(state, (transaction) => {
        next = state.apply(transaction);
      });
      return next ?? panic("The legacy toggle did not apply.");
    };
    for (const [markName, commandName] of toggles) {
      const doc = schema.node("doc", null, [
        schema.node("paragraph", { paraId: "12345678" }, [
          schema.text("ab", [schema.mark(markName)]),
          schema.text("cd"),
        ]),
      ]);
      const session = createCanonicalSession({
        package: { document: { content: proseDocToBlocks(doc) } },
      }).unwrap();
      const initial = EditorState.create({
        schema,
        doc: session.projection.doc,
        selection: TextSelection.create(session.projection.doc, 1, 5),
      });
      const marked = (value: EditorState) => {
        const flags: boolean[] = [];
        value.doc.descendants((node) => {
          if (!node.isText) return;
          // Underline turns off as the `none` style, so only a visible style counts.
          const has = node.marks.some(
            (mark) =>
              mark.type.name === markName && mark.attrs["style"] !== "none",
          );
          for (const _ of node.text ?? "") flags.push(has);
        });
        return flags;
      };
      let legacy = initial;
      let canonical = initial;
      for (const expected of [false, true]) {
        const command = singletonManager.requireCommand(commandName)();
        legacy = runLegacy(legacy, command);
        const intents =
          getCanonicalCommandIntents(command, canonical) ??
          (() => {
            throw new TypeError(`${commandName} has no canonical descriptor.`);
          })();
        canonical = publishCanonicalProjection({
          session,
          state: canonical,
          commit: prepareCanonicalCommands(
            session,
            canonical,
            intents,
          ).unwrap(),
        }).unwrap().state;
        expect(marked(legacy)).toEqual([
          expected,
          expected,
          expected,
          expected,
        ]);
        expect(marked(canonical)).toEqual([
          expected,
          expected,
          expected,
          expected,
        ]);
      }
    }
  });
});
