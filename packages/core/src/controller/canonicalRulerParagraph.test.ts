import { describe, expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import {
  applyRulerParagraphEdit,
  rulerParagraphCommand,
  type RulerParagraphEdit,
} from "../prosemirror/commands/rulerParagraph";
import type { CanonicalGap } from "../types/canonicalCapabilities";
import type { Document, ParagraphFormatting } from "../types/document";
import {
  createHiddenEditorManager,
  type HiddenEditorManagerDeps,
} from "./hiddenEditorManager";

/** One paragraph with a left tab stop at one inch. */
const source = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          formatting: { tabs: [{ position: 1440, alignment: "left" }] },
          content: [
            { type: "run", content: [{ type: "text", text: "Alpha Beta" }] },
          ],
        },
      ],
    },
  },
});

const deps = (
  host: HTMLElement,
  document: Document,
  refused: (CanonicalGap | undefined)[],
) =>
  ({
    getHost: () => host,
    getDocument: () => document,
    getStyles: () => null,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: () => false,
    getExperimentalSession: () => "canonical",
    getDocumentIdentity: () => "ruler-fixture",
    getDocumentContext: () => document,
    onTransaction: () => undefined,
    onCanonicalCompositionEnd: () => undefined,
    onSelectionChange: () => undefined,
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => undefined,
    onEditorViewReady: () => undefined,
    onEditorViewDestroy: () => undefined,
    onRemoteSelectionsChange: () => undefined,
    onSessionRefusal: (_message, ground) => refused.push(ground.gap),
  }) satisfies HiddenEditorManagerDeps;

type Run = "owner" | "raw";

/**
 * Mount a canonical editor, make one ruler edit and return the paragraph's
 * formatting in the session's document, with the refusals it reported.
 * `raw` runs the same command against the view's own state and dispatch,
 * as the adapters did before the edit went through the command owner.
 */
const edited = (edit: RulerParagraphEdit, run: Run) => {
  const releaseDom = acquireHarnessDom();
  const host = document.createElement("div");
  document.body.append(host);
  const refused: (CanonicalGap | undefined)[] = [];
  const manager = createHiddenEditorManager(deps(host, source(), refused));
  try {
    manager.ensureView();
    const view = manager.getView();
    if (!view) throw new Error("The canonical view did not mount.");
    if (run === "owner") expect(applyRulerParagraphEdit(view, edit)).toBe(true);
    else rulerParagraphCommand(edit)(view.state, view.dispatch, view);
    const [paragraph] =
      manager.api.getCanonicalDocument()?.package.document.content ?? [];
    const formatting: ParagraphFormatting | undefined =
      paragraph?.type === "paragraph" ? paragraph.formatting : undefined;
    return { formatting, canUndo: manager.api.canUndo(), refused };
  } finally {
    manager.destroyView();
    host.remove();
    releaseDom();
  }
};

describe("ruler paragraph edits in a canonical session", () => {
  test("indents and tab stop removal are journalled paragraph formatting", () => {
    const left = edited({ type: "indentLeft", twips: 720 }, "owner");
    expect(left.formatting?.indentLeft).toBe(720);
    expect(left.canUndo).toBe(true);
    expect(left.refused).toEqual([]);

    const right = edited({ type: "indentRight", twips: 360 }, "owner");
    expect(right.formatting?.indentRight).toBe(360);

    const first = edited({ type: "indentFirstLine", twips: 240 }, "owner");
    expect(first.formatting?.indentFirstLine).toBe(240);
    expect(first.formatting?.hangingIndent ?? false).toBe(false);

    const hanging = edited({ type: "indentFirstLine", twips: -240 }, "owner");
    expect(hanging.formatting?.indentFirstLine).toBe(240);
    expect(hanging.formatting?.hangingIndent).toBe(true);

    const tab = edited({ type: "removeTabStop", position: 1440 }, "owner");
    expect(tab.formatting?.tabs ?? []).toEqual([]);
    expect(tab.canUndo).toBe(true);
    expect(tab.refused).toEqual([]);
  });

  test("the same edit dispatched on the view leaves the session's document as it was", () => {
    // The check above fails for a ruler that dispatches on the view: the
    // session never sees the edit.
    const left = edited({ type: "indentLeft", twips: 720 }, "raw");
    expect(left.formatting?.indentLeft).toBeUndefined();
    expect(left.canUndo).toBe(false);
    const tab = edited({ type: "removeTabStop", position: 1440 }, "raw");
    expect(tab.formatting?.tabs).toEqual([
      { position: 1440, alignment: "left" },
    ]);
    expect(tab.canUndo).toBe(false);
  });
});
