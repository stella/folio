import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, setSystemTime, spyOn, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { EditorState, TextSelection } from "prosemirror-state";
import { IntlProvider } from "use-intl";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { createCanonicalSession } from "@stll/folio-core/controller/canonicalSession";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { schema } from "@stll/folio-core/prosemirror/schema";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { assertExactModel } from "../../../../test/exactModel";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Comment, DrawingContent } from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";
import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const createErrorObserver = () => {
  const errors: Error[] = [];
  return { errors, onError: (error: Error) => errors.push(error) };
};

const recordDocumentNotifications = (notifications: Document[]) => (next: Document) => {
  notifications.push(next);
};

const captureDocumentNotifications = () => {
  const pending = new Map<number, () => void>();
  const schedule = window.setTimeout.bind(window);
  const clear = window.clearTimeout.bind(window);
  const scheduleSpy = spyOn(window, "setTimeout").mockImplementation((callback, delay, ...args) => {
    if (delay !== 250 || typeof callback !== "function") return schedule(callback, delay, ...args);
    const handle = schedule(() => {}, 60_000);
    pending.set(handle, () => callback(...args));
    return handle;
  });
  const clearSpy = spyOn(window, "clearTimeout").mockImplementation((handle) => {
    if (handle !== undefined) pending.delete(handle);
    clear(handle);
  });
  return {
    get pendingCount() {
      return pending.size;
    },
    fire: () => {
      for (const [handle, callback] of [...pending]) {
        pending.delete(handle);
        clear(handle);
        callback();
      }
    },
    restore: () => {
      for (const handle of pending.keys()) clear(handle);
      scheduleSpy.mockRestore();
      clearSpy.mockRestore();
    },
  };
};

const SHAPES = [
  "header-footer",
  "mixed-lists",
  "single-decimal-list",
  "image",
  "image-leading",
  "image-trailing",
] as const;
const MODES = ["editing", "suggesting"] as const;
const COMPLETIONS = ["cancel", "commit"] as const;

const createShapeBuffer = async (shape: (typeof SHAPES)[number]) => {
  if (shape === "mixed-lists" || shape === "single-decimal-list") {
    const markdown =
      shape === "mixed-lists"
        ? "Intro paragraph.\n\n1. Alpha\n2. Beta\n\nPlain text.\n\n- Gamma\n- Delta\n\nTail."
        : "Intro paragraph.\n\n1. Alpha\n2. Beta\n\nPlain text.\n\nTail.";
    return createDocx(fromMarkdown(markdown));
  }

  const document = createEmptyDocument({ initialText: "Body paragraph under a header." });
  if (shape === "header-footer") {
    document.package.headers = new Map([
      [
        "rIdHeader1",
        {
          type: "header",
          hdrFtrType: "default",
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "Header text" }] }],
            },
          ],
        },
      ],
    ]);
    document.package.footers = new Map([
      [
        "rIdFooter1",
        {
          type: "footer",
          hdrFtrType: "default",
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "Footer text" }] }],
            },
          ],
        },
      ],
    ]);
    document.package.document.finalSectionProperties = {
      ...document.package.document.finalSectionProperties,
      headerReferences: [{ type: "default", rId: "rIdHeader1" }],
      footerReferences: [{ type: "default", rId: "rIdFooter1" }],
    };
  } else {
    const drawing = {
      type: "drawing",
      image: {
        type: "image",
        src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
        mimeType: "image/png",
        size: { width: 9_525, height: 9_525 },
        wrap: { type: "inline" },
      },
    } satisfies DrawingContent;
    document.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            content: [
              ...(shape === "image-leading"
                ? []
                : [{ type: "text" as const, text: "Text before the picture" }]),
              drawing,
              ...(shape === "image-trailing"
                ? []
                : [{ type: "text" as const, text: " and after it." }]),
            ],
          },
        ],
      },
    ];
  }
  return createDocx(document);
};

for (const shape of SHAPES) {
  for (const mode of MODES) {
    for (const completion of COMPLETIONS) {
      test(
        `canonical comments render during ${shape} ${mode} composition ${completion}`,
        async () => {
          await assertProperty(
            fc.asyncProperty(
              fc.constantFrom("!", "?", "Ω"),
              fc.constantFrom("beforeSelection", "beforeRender", "afterRender"),
              async (preEdit, timerPhase) => {
                const container = document.createElement("div");
                document.body.append(container);
                const root = createRoot(container);
                const editor = createRef<DocxEditorRef>();
                const bytes = await createShapeBuffer(shape);
                const messages = getFolioMessages("en");
                const observer = createErrorObserver();
                const notifications: Document[] = [];
                const onDocumentChange = recordDocumentNotifications(notifications);
                const notificationTimer = captureDocumentNotifications();
                const controls: { comments: Comment[]; showToolbar: boolean } = {
                  comments: [],
                  showToolbar: false,
                };
                const renderEditor = () => (
                  <IntlProvider locale="en" timeZone="UTC" messages={messages}>
                    <DocxEditor
                      ref={editor}
                      documentBuffer={bytes}
                      experimentalSession="canonical"
                      comments={controls.comments}
                      onError={observer.onError}
                      onChange={onDocumentChange}
                      showToolbar={controls.showToolbar}
                    />
                  </IntlProvider>
                );
                try {
                  await act(async () => root.render(renderEditor()));
                  await act(async () => editor.current?.loadDocumentBuffer(bytes));
                  await act(async () => editor.current?.ensureEditorView({ focus: false }));
                  const api = editor.current?.getEditor() ?? panic("Expected canonical editor");
                  const view = api.getView() ?? panic("Expected mounted canonical view");
                  const sessionMode =
                    mode === "editing"
                      ? { type: "editing" as const }
                      : { type: "suggesting" as const, author: "Composition author" };
                  expect(api.setCanonicalMode(sessionMode)).toBe(true);
                  let from: number | undefined;
                  view.state.doc.descendants((node, pos) => {
                    if (from === undefined && node.isText && node.nodeSize >= 3) from = pos;
                  });
                  const compositionFrom = from ?? panic("Expected a text span beside the picture");
                  const compositionTo = compositionFrom + 3;
                  await act(async () => notificationTimer.fire());
                  notifications.length = 0;
                  await act(async () => {
                    view.dispatch(
                      view.state.tr.setSelection(
                        TextSelection.create(view.state.doc, compositionTo),
                      ),
                    );
                    view.dom.dispatchEvent(
                      new InputEvent("beforeinput", {
                        bubbles: true,
                        cancelable: true,
                        inputType: "insertText",
                        data: preEdit,
                      }),
                    );
                    view.dispatch(
                      view.state.tr.setSelection(
                        TextSelection.create(view.state.doc, compositionFrom, compositionTo),
                      ),
                    );
                  });
                  expect(notificationTimer.pendingCount).toBeGreaterThan(0);
                  const baseline = api.getCanonicalDocument() ?? panic("Expected committed model");
                  const baselineProjection = view.state.doc;
                  setSystemTime(new Date("2026-10-05T12:00:00Z"));
                  const expected = createCanonicalSession(baseline).unwrap();
                  expected.setMode(sessionMode);
                  const expectedState = EditorState.create({
                    schema,
                    doc: expected.projection.doc,
                  });
                  if (completion === "commit") {
                    expected
                      .prepareReplace(expectedState, {
                        from: compositionFrom,
                        to: compositionTo,
                        text: "alpha",
                        semantic: "composition",
                      })
                      .unwrap()
                      .publish()
                      .unwrap();
                  }
                  await act(async () => {
                    view.dom.dispatchEvent(
                      new CompositionEvent("compositionstart", { bubbles: true }),
                    );
                    view.dispatch(
                      view.state.tr
                        .insertText("alpha", compositionFrom, compositionTo)
                        .setMeta("composition", 1),
                    );
                  });
                  const fireDuringComposition = async () => {
                    await act(async () => expect(notificationTimer.fire).not.toThrow());
                    expect(notifications).toEqual([]);
                  };
                  if (timerPhase === "beforeSelection") await fireDuringComposition();
                  expect(api.getCanonicalDocument).toThrow(
                    "Composition must finish before taking a snapshot.",
                  );
                  // Selection notifications and a controlled-prop render must read only
                  // published comments, while public snapshots retain their refusal.
                  await act(async () => {
                    view.dispatch(
                      view.state.tr.setSelection(
                        TextSelection.create(view.state.doc, compositionFrom, compositionFrom + 1),
                      ),
                    );
                  });
                  if (timerPhase === "beforeRender") await fireDuringComposition();
                  controls.comments = [...controls.comments];
                  controls.showToolbar = true;
                  await act(async () => root.render(renderEditor()));
                  if (timerPhase === "afterRender") await fireDuringComposition();
                  expect(notifications).toEqual([]);
                  expect(editor.current?.getEditor()?.getView()).toBe(view);
                  expect(view.isDestroyed).toBe(false);
                  expect(api.getCanonicalDocument).toThrow(
                    "Composition must finish before taking a snapshot.",
                  );
                  await act(async () => {
                    if (completion === "cancel") {
                      // Native cancellation restores the selected source, including
                      // its run marks, before the final flush.
                      view.dispatch(
                        view.state.tr
                          .replaceWith(
                            compositionFrom,
                            compositionFrom + "alpha".length,
                            baselineProjection.slice(compositionFrom, compositionTo).content,
                          )
                          .setMeta("composition", 1),
                      );
                      expect(view.state.doc.eq(baselineProjection)).toBe(true);
                    }
                    view.dom.dispatchEvent(
                      new CompositionEvent("compositionend", { bubbles: true }),
                    );
                    await new Promise<void>((resolve) => setTimeout(resolve, 40));
                  });
                  const committed =
                    api.getCanonicalDocument() ?? panic("Expected completed snapshot");
                  assertExactModel(committed, expected.document);
                  await act(async () => notificationTimer.fire());
                  expect(notifications).toHaveLength(1);
                  assertExactModel(
                    notifications.at(0) ?? panic("Expected final notification"),
                    expected.document,
                  );
                  await act(async () => notificationTimer.fire());
                  expect(notifications).toHaveLength(1);
                  expect(editor.current?.getEditor()?.getView()).toBe(view);
                  expect(observer.errors).toEqual([]);
                } finally {
                  setSystemTime();
                  await act(async () => root.unmount());
                  notificationTimer.restore();
                  container.remove();
                }
              },
            ),
            { numRuns: 3 },
          );
        },
        propertyTestTimeout(30_000),
      );
    }
  }
}
