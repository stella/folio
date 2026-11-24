import { describe, expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import {
  applyRulerParagraphEdit,
  rulerParagraphCommand,
  type RulerParagraphEdit,
} from "../prosemirror/commands/rulerParagraph";
import {
  CANONICAL_GAP,
  type CanonicalGap,
} from "../types/canonicalCapabilities";
import type {
  Document,
  Paragraph,
  ParagraphFormatting,
} from "../types/document";
import { createHeaderFooterEditorManager } from "./headerFooterEditorManager";
import {
  createHiddenEditorManager,
  type HiddenEditorManagerDeps,
} from "./hiddenEditorManager";
import { createNoteEditorManager } from "./noteEditorManager";

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
    getEmbeddedFontFamilyMap: () => undefined,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getReadOnly: () => false,
    getDocumentIdentity: () => "ruler-fixture",
    getDocumentContext: () => document,
    onTransaction: () => undefined,
    onCanonicalCompositionLayout: () => undefined,
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
  test.each(["header", "footer", "footnote", "endnote"] as const)(
    "%s ruler edits refuse without changing the story or body",
    (kind) => {
      const releaseDom = acquireHarnessDom();
      const bodyHost = document.createElement("div");
      const storyHost = document.createElement("div");
      document.body.append(bodyHost, storyHost);
      const documentSource = source();
      const storyContent = (paraId: string): Paragraph[] => [
        {
          type: "paragraph",
          paraId,
          formatting: { tabs: [{ position: 1440, alignment: "left" }] },
          content: [
            { type: "run", content: [{ type: "text", text: "Story" }] },
          ],
        },
      ];
      documentSource.package.headers = new Map([
        [
          "rIdHeader",
          {
            type: "header",
            hdrFtrType: "default",
            content: storyContent("23456789"),
          },
        ],
      ]);
      documentSource.package.footers = new Map([
        [
          "rIdFooter",
          {
            type: "footer",
            hdrFtrType: "default",
            content: storyContent("34567890"),
          },
        ],
      ]);
      documentSource.package.footnotes = [
        { type: "footnote", id: 1, content: storyContent("45678901") },
      ];
      documentSource.package.endnotes = [
        { type: "endnote", id: 1, content: storyContent("56789012") },
      ];
      const refused: (CanonicalGap | undefined)[] = [];
      const manager = createHiddenEditorManager(
        deps(bodyHost, documentSource, refused),
      );
      const storyDeps = {
        getHost: () => storyHost,
        getDocument: () => documentSource,
        getCanonicalApi: () => manager.api,
        readCanonicalDocument: () => manager.readCanonicalDocument(),
        getStyles: () => null,
        getTheme: () => null,
        onSessionRefusal: deps(bodyHost, documentSource, refused)
          .onSessionRefusal,
      };
      const parts = createHeaderFooterEditorManager(storyDeps);
      const notes = createNoteEditorManager(storyDeps);
      try {
        manager.ensureView();
        parts.sync();
        const view =
          kind === "header" || kind === "footer"
            ? parts.getView(kind === "header" ? "rIdHeader" : "rIdFooter")
            : notes.activate({ kind, noteId: 1 });
        if (!view) throw new Error("The canonical story view did not mount.");
        expect(view.state.doc.textContent).toBe("Story");
        const before = manager.api.getCanonicalDocument();
        const projectedBefore = view.state.doc;
        const edits = [
          { type: "indentLeft", twips: 720 },
          { type: "indentRight", twips: 360 },
          { type: "indentFirstLine", twips: -240 },
          { type: "removeTabStop", position: 1440 },
        ] as const satisfies readonly RulerParagraphEdit[];
        for (const edit of edits) {
          applyRulerParagraphEdit(view, edit);
          expect(view.state.doc.eq(projectedBefore)).toBe(true);
          expect(manager.api.getCanonicalDocument()).toEqual(before);
          expect(manager.api.canUndo()).toBe(false);
        }
        // The native input boundary reports a refusal once until the next gesture.
        expect(refused).toEqual([CANONICAL_GAP.dispatch]);
      } finally {
        parts.destroy();
        notes.destroy();
        manager.destroyView();
        bodyHost.remove();
        storyHost.remove();
        releaseDom();
      }
    },
  );

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
