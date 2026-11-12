import { GlobalRegistrator } from "@happy-dom/global-registrator";
import {
  afterAll,
  afterEach,
  beforeAll,
  describe,
  expect,
  test,
} from "bun:test";
import { history, undoDepth } from "prosemirror-history";
import { EditorState, Plugin } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "@stll/folio-core/controller/canonicalSession";
import { schema } from "@stll/folio-core/prosemirror/schema";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { createHistoryBridgePlugin } from "./historyBridge";
import type { HistoryChange } from "./historyBridge";

const fail = (message: string): never => {
  throw new Error(message);
};

/** Longer than the session's undo grouping window. */
const PAUSE_MS = 1_000;

const seed = (): Document => {
  const document = createEmptyDocument({ initialText: "Start" });
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") fail("Expected a paragraph.");
  paragraph.paraId = "12345678";
  return document;
};

/**
 * What the bridge did before the canonical session: a new step whenever
 * ProseMirror's own history grew. The session keeps its commits out of that
 * history, so this one never sees a step.
 */
const createProseMirrorHistoryBridge = (
  onChange: (change: HistoryChange) => void,
): Plugin =>
  new Plugin({
    view: () => ({
      update: (view, previous) => {
        if (view.state.doc === previous.doc) return;
        onChange(
          undoDepth(view.state) === undoDepth(previous)
            ? "sameStep"
            : "newStep",
        );
      },
    }),
  });

const views: EditorView[] = [];

const createEditor = (bridge = createHistoryBridgePlugin) => {
  const session = createCanonicalSession(seed()).unwrap();
  const changes: HistoryChange[] = [];
  const plugins = [history(), bridge((change) => changes.push(change))];
  const freshState = () =>
    EditorState.create({ schema, doc: session.projection.doc, plugins });
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, { state: freshState() });
  views.push(view);
  // As the editor publishes: the session accepts the projection, then the
  // view takes the state it staged.
  const publish = (commit: CanonicalCommit) =>
    view.updateState(
      publishCanonicalProjection({
        state: view.state,
        commit,
        session,
      }).unwrap().state,
    );
  let time = 0;

  return {
    session,
    view,
    changes,
    /** Type at the caret, after a pause or while typing on. */
    type: (text: string, { pause }: { pause: boolean }) => {
      time += pause ? PAUSE_MS : 1;
      const at = view.state.selection.head;
      publish(
        session
          .prepareReplace(view.state, {
            from: at,
            to: at,
            text,
            semantic: "typing",
            time,
          })
          .unwrap(),
      );
    },
    undo: () => publish(session.prepareUndo(view.state).unwrap()),
    redo: () => publish(session.prepareRedo(view.state).unwrap()),
    reload: () => view.updateState(freshState()),
  };
};

/** Type three undo steps, the second one in two keystrokes. */
const typeThreeSteps = (editor: ReturnType<typeof createEditor>) => {
  editor.type("a", { pause: true });
  editor.type("b", { pause: true });
  editor.type("c", { pause: false });
  editor.type("d", { pause: true });
};

beforeAll(() => GlobalRegistrator.register());

afterEach(() => {
  for (const view of views.splice(0)) {
    const mount = view.dom.parentElement;
    view.destroy();
    mount?.remove();
  }
});

afterAll(() => GlobalRegistrator.unregister());

describe("createHistoryBridgePlugin", () => {
  test("reports one new step per undo step the session opens, and undo / redo apart", () => {
    const editor = createEditor();
    typeThreeSteps(editor);
    editor.undo();
    editor.redo();

    expect(editor.changes).toEqual([
      "newStep",
      "newStep",
      "sameStep",
      "newStep",
      "undoRedo",
      "undoRedo",
    ]);
  });

  test("the host's stack walks the session's steps one for one", () => {
    const editor = createEditor();
    typeThreeSteps(editor);
    const steps = editor.changes.filter((change) => change === "newStep");

    let undone = 0;
    while (editor.session.canUndo) {
      editor.undo();
      undone += 1;
    }
    expect(undone).toBe(steps.length);
    expect(editor.view.state.doc.textContent).toBe("Start");
  });

  test("a bridge that reads ProseMirror's history sees none of the session's steps", () => {
    const bridged = createEditor();
    const mirrored = createEditor(createProseMirrorHistoryBridge);
    typeThreeSteps(bridged);
    typeThreeSteps(mirrored);

    expect(mirrored.changes).not.toContain("newStep");
    expect(mirrored.changes).not.toEqual(bridged.changes);
  });

  test("a reload is not an edit", () => {
    const editor = createEditor();
    editor.type("a", { pause: true });
    editor.reload();

    expect(editor.changes).toEqual(["newStep"]);
  });
});
