import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { resetCanvasContext } from "@stll/folio-core/layout-engine/measure/measureContainer";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "@stll/folio-core/layout-engine/measure/measureProvider";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const createErrorRecorder = (errors: Error[]) => (error: Error) =>
  errors.push(error);

test("read-only import repairs duplicate bookmarks and renders their first reference target", async () => {
  const bytes = await Bun.file(
    new URL(
      "../../../core/src/docx/__tests__/__fixtures__/corpus/duplicate-bookmark-ids.docx",
      import.meta.url,
    ),
  ).arrayBuffer();
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const errors: Error[] = [];
  const onError = createErrorRecorder(errors);
  const installedProvider = getMeasureProvider();
  const installedInstrumentation = globalThis.__folioLayoutInstrumentation;
  const layoutErrors: string[] = [];
  globalThis.__folioLayoutInstrumentation = {
    onMeasureBlockError: ({ message }) => layoutErrors.push(message),
    onLayoutError: ({ message }) => layoutErrors.push(message),
  };
  // The composition root installs canvas measurement on each layout run.
  // Stub the missing HappyDOM canvas, keeping the document pipeline real.
  const canvasDescriptor = Object.getOwnPropertyDescriptor(
    HTMLCanvasElement.prototype,
    "getContext",
  );
  Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
    configurable: true,
    value: () => ({
      font: "",
      measureText: (text: string) => ({
        width: text.length * 7,
        actualBoundingBoxAscent: 12,
        actualBoundingBoxDescent: 4,
        fontBoundingBoxAscent: 12,
        fontBoundingBoxDescent: 4,
      }),
    }),
  });
  resetCanvasContext();
  try {
    await act(async () => {
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <DocxEditor
            accessibleName="Document editor"
            ref={editor}
            documentBuffer={bytes}
            readOnly
            showToolbar={false}
            onError={onError}
          />
        </IntlProvider>,
      );
    });
    await act(async () => editor.current?.loadDocumentBuffer(bytes));
    await act(async () => editor.current?.ensureEditorView({ focus: false }));
    for (
      let attempt = 0;
      attempt < 400 &&
      !container
        .querySelector(".layout-page")
        ?.textContent?.includes("Later target");
      attempt++
    ) {
      await act(
        async () => await new Promise((resolve) => setTimeout(resolve, 10)),
      );
    }
    expect(errors).toEqual([]);
    expect(layoutErrors).toEqual([]);
    expect(editor.current?.getDocument()).toBeTruthy();
    const page = container.querySelector(".layout-page");
    expect(page).not.toBeNull();
    expect(page?.innerHTML).toContain("First target");
    expect(page?.textContent).toContain("Later target");
    expect(page?.textContent?.match(/First target/g)).toHaveLength(2);
    expect(page?.textContent).not.toContain("Stale reference");
  } finally {
    await act(async () => root.unmount());
    container.remove();
    setMeasureProvider(installedProvider);
    if (canvasDescriptor)
      Object.defineProperty(
        HTMLCanvasElement.prototype,
        "getContext",
        canvasDescriptor,
      );
    resetCanvasContext();
    globalThis.__folioLayoutInstrumentation = installedInstrumentation;
  }
}, 15_000);
