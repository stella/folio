import { afterAll, beforeAll, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { DOMParser } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import fc from "fast-check";
import { panic } from "better-result";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { shapeArrayBuffer } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  createHarnessPlugins,
  HARNESS_AUTHOR,
  HeadlessEditorView,
  parseShapeDocument,
  resolveAllChanges,
  summarizeState,
} from "../../__tests__/editorHarness";
import { BROWSER_NOTES_SEED_197 } from "../../../../../tests/visual/browserInputRegressions";

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

test("browser notes seed 197 paste and caret deletion replay", async () => {
  const regression = BROWSER_NOTES_SEED_197;
  const base = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer(regression.trace.shape)),
  );
  const states = [];
  for (const mode of ["editing", "suggesting"] as const) {
    const view = new HeadlessEditorView(createHarnessState(base, mode));
    const paste = (source: string) => {
      let html = source;
      const pasteView = new EditorView(document.createElement("div"), { state: view.state });
      try {
        for (const plugin of pasteView.state.plugins) {
          const transform = plugin.props.transformPastedHTML;
          if (transform) html = transform.call(plugin, html, pasteView);
        }
      } finally {
        pasteView.destroy();
      }
      const host = document.createElement("div");
      host.innerHTML = html;
      view.paste(DOMParser.fromSchema(view.state.schema).parseSlice(host));
    };
    for (const action of regression.trace.actions) {
      switch (action.kind) {
        case "pasteWordHtml":
          paste(action.html);
          expect(view.state.doc.textContent).toContain("Closing\u00a0");
          break;
        case "pasteTable":
          paste(action.html);
          break;
        case "cut":
          // Native cut at a caret does not change the document.
          expect(view.state.selection.empty).toBe(true);
          break;
        case "delete":
          expect(view.pressKey("Delete")).toBe(true);
          break;
        case "backspace":
          // Positive control: the second key must cross the first key's
          // retained deletion before it can retract the pasted space.
          if (mode === "suggesting")
            expect(
              view.state.selection.$from.nodeBefore?.marks.some(
                (mark) => mark.type.name === "deletion",
              ),
            ).toBe(true);
          expect(view.pressKey("Backspace")).toBe(true);
          break;
        case "selectionDrag": {
          // Reproduce the drag's semantic range; browser replay proves the gesture.
          let noteRange: { from: number; to: number } | undefined;
          view.state.doc.descendants((node, position) => {
            if (noteRange || !node.marks.some((mark) => mark.type.name === "footnoteRef")) return;
            noteRange = { from: position, to: position + node.nodeSize };
          });
          if (!noteRange) panic("The standing trace lost its note drag target.");
          view.dispatch(
            view.state.tr.setSelection(
              TextSelection.create(view.state.doc, noteRange.from, noteRange.to),
            ),
          );
          break;
        }
        default: {
          const unhandled: never = action;
          panic(`Unhandled standing action ${String(unhandled)}.`);
        }
      }
    }
    states.push(summarizeState(resolveAllChanges(view.state, "accept")));
  }
  expect(states.at(1)).toEqual(states.at(0));
});

test(
  "caret direction reversal deletes the adjacent visible unit across retained revisions and markers",
  async () => {
    const base = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("notes")));
    const { schema } = createHarnessState(base, "editing");
    const check = ({
      direction,
      deletedLength,
      character,
      ownership,
      markers,
      deletedOwner,
    }: {
      direction: "forward" | "backward";
      deletedLength: number;
      character: string;
      ownership: "own" | "other" | "existing";
      markers: boolean;
      deletedOwner: "own" | "other";
    }) => {
      const revision = { revisionId: 100, author: HARNESS_AUTHOR, date: "2026-01-01T00:00:00Z" };
      const insertion =
        ownership === "existing"
          ? []
          : [
              schema.mark("insertion", {
                ...revision,
                author: ownership === "own" ? HARNESS_AUTHOR : "Other",
                revisionId: 101,
              }),
            ];
      const priorDeletion = schema.mark("deletion", {
        ...revision,
        author: deletedOwner === "own" ? HARNESS_AUTHOR : "Other",
      });
      const hidden = schema.text("x".repeat(deletedLength), [priorDeletion]);
      const boundaries = markers
        ? [
            schema.node("bookmarkBoundary", { type: "start", id: 1, name: "range" }),
            hidden,
            schema.node("bookmarkBoundary", { type: "end", id: 1 }),
          ]
        : [hidden];
      const content =
        direction === "forward"
          ? [schema.text("A"), schema.text(character, insertion), ...boundaries, schema.text("RZ")]
          : [schema.text("AL"), ...boundaries, schema.text(character, insertion), schema.text("Z")];
      const doc = schema.node("doc", null, [
        schema.node("paragraph", { paraId: "12345678" }, content),
      ]);
      const caret = direction === "forward" ? 2 + character.length : 3;
      const state = EditorState.create({
        doc,
        selection: TextSelection.create(doc, caret),
        plugins: createHarnessPlugins(base, "suggesting"),
      });
      const tracked = new HeadlessEditorView(state);
      expect(tracked.state.doc.textContent).toContain("x".repeat(deletedLength));
      const editedDoc = schema.node("doc", null, [
        schema.node("paragraph", { paraId: "12345678" }, [
          schema.text(direction === "forward" ? `A${character}RZ` : `AL${character}Z`),
        ]),
      ]);
      const edited = new HeadlessEditorView(
        EditorState.create({
          doc: editedDoc,
          selection: TextSelection.create(editedDoc, caret),
          plugins: createHarnessPlugins(base, "editing"),
        }),
      );
      const keys = direction === "forward" ? ["Delete", "Backspace"] : ["Backspace", "Delete"];
      for (const key of keys) {
        expect(tracked.pressKey(key)).toBe(true);
        expect(edited.pressKey(key)).toBe(true);
      }
      const accepted = resolveAllChanges(tracked.state, "accept");
      expect(
        accepted.doc.textContent,
        JSON.stringify({ direction, ownership, deletedOwner, markers, deletedLength, character }),
      ).toBe(edited.state.doc.textContent);
      expect(accepted.doc.textContent).toBe("AZ");
      const rejected = resolveAllChanges(tracked.state, "reject");
      const originalCharacter = ownership === "existing" ? character : "";
      expect(rejected.doc.textContent).toBe(
        direction === "forward"
          ? `A${originalCharacter}${"x".repeat(deletedLength)}RZ`
          : `AL${"x".repeat(deletedLength)}${originalCharacter}Z`,
      );
      expect(tracked.state.doc.textContent.includes(character)).toBe(ownership !== "own");
      const retainedMarkers: string[] = [];
      let retainedDeletedText = "";
      let retainedForeignInsertion = "";
      const foreignInsertion = ownership === "other" ? insertion.at(0) : undefined;
      tracked.state.doc.descendants((node) => {
        if (node.type.name === "bookmarkBoundary") retainedMarkers.push(node.attrs["type"]);
        if (node.isText && node.marks.some((mark) => mark.eq(priorDeletion)))
          retainedDeletedText += node.text;
        if (node.isText && foreignInsertion && node.marks.some((mark) => mark.eq(foreignInsertion)))
          retainedForeignInsertion += node.text;
      });
      expect(retainedMarkers).toEqual(markers ? ["start", "end"] : []);
      expect(retainedDeletedText).toContain("x".repeat(deletedLength));
      expect(retainedForeignInsertion).toBe(ownership === "other" ? character : "");
    };
    // Every direction and ownership gets the full marker/run-length boundary matrix.
    for (const direction of ["forward", "backward"] as const)
      for (const ownership of ["own", "other", "existing"] as const)
        for (const deletedOwner of ["own", "other"] as const)
          for (const markers of [false, true])
            for (let deletedLength = 1; deletedLength <= 8; deletedLength++)
              for (const character of ["\u00a0", "Q"])
                check({ direction, ownership, deletedOwner, markers, deletedLength, character });
    assertProperty(
      fc.property(
        fc.record({
          direction: fc.constantFrom("forward", "backward"),
          ownership: fc.constantFrom("own", "other", "existing"),
          deletedOwner: fc.constantFrom("own", "other"),
          markers: fc.boolean(),
          deletedLength: fc.integer({ min: 1, max: 64 }),
          character: fc.constantFrom("\u00a0", "Q", " "),
        }),
        check,
      ),
      {
        seed: 197,
        numRuns: 100,
        id: "caret direction reversal deletes the adjacent visible unit across retained revisions and markers",
      },
    );
  },
  propertyTestTimeout(10_000),
);
