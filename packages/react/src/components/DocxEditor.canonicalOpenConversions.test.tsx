import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic, Result } from "better-result";
import { afterAll, expect, spyOn, test } from "bun:test";
import { IntlProvider } from "use-intl";

import * as canonicalSessions from "@stll/folio-core/controller/canonicalSession";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { resetCanvasContext } from "@stll/folio-core/layout-engine/measure/measureContainer";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "@stll/folio-core/layout-engine/measure/measureProvider";
import * as conversion from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { CANONICAL_GAP } from "@stll/folio-core/types/canonicalCapabilities";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

// Converting a document to ProseMirror takes time in proportion to the
// document. A canonical session seeds its view with its own projection, so a
// canonical open converts the body once; the pages convert it themselves only
// when no view holds it, as after a refused seed.

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

/** HappyDOM has no canvas: measure text with a fixed width per character. */
const stubCanvas = () => {
  const installedProvider = getMeasureProvider();
  const descriptor = Object.getOwnPropertyDescriptor(
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
  return () => {
    setMeasureProvider(installedProvider);
    if (descriptor)
      Object.defineProperty(
        HTMLCanvasElement.prototype,
        "getContext",
        descriptor,
      );
    resetCanvasContext();
  };
};

const mountCanonicalEditor = async (text: string, errors: Error[]) => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const onError = createErrorRecorder(errors);
  await act(async () => {
    root.render(
      <IntlProvider
        locale="en"
        timeZone="UTC"
        messages={getFolioMessages("en")}
      >
        <DocxEditor
          ref={editor}
          accessibleName="Document editor"
          document={createEmptyDocument({ initialText: text })}
          experimentalSession="canonical"
          showToolbar={false}
          onError={onError}
        />
      </IntlProvider>,
    );
  });
  return {
    container,
    ref: () => editor.current ?? panic("Expected a mounted editor"),
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
};

/** Waits until the painted pages show `text`. */
const waitForPaintedText = async (container: HTMLElement, text: string) => {
  for (
    let attempt = 0;
    attempt < 400 &&
    !container.querySelector(".layout-page")?.textContent?.includes(text);
    attempt++
  ) {
    await act(
      async () => await new Promise((resolve) => setTimeout(resolve, 10)),
    );
  }
  expect(container.querySelector(".layout-page")?.textContent).toContain(text);
};

test("a canonical open converts the document to ProseMirror once", async () => {
  const restoreCanvas = stubCanvas();
  const conversions = spyOn(conversion, "toProseDoc");
  const errors: Error[] = [];
  const mounted = await mountCanonicalEditor("First open", errors);
  try {
    await waitForPaintedText(mounted.container, "First open");
    expect(mounted.ref().getEditor()?.getView()).not.toBeNull();
    expect(conversions).toHaveBeenCalledTimes(1);

    // A document loaded into the open editor: the view is seeded again, and
    // the pages lay out the view's state instead of converting their own.
    const bytes = await createDocx(
      createEmptyDocument({ initialText: "Second open" }),
    );
    conversions.mockClear();
    await act(async () => mounted.ref().loadDocumentBuffer(bytes));
    await waitForPaintedText(mounted.container, "Second open");
    expect(mounted.ref().getEditor()?.getView()).not.toBeNull();
    expect(conversions).toHaveBeenCalledTimes(1);
    expect(errors).toEqual([]);
  } finally {
    await mounted.unmount();
    conversions.mockRestore();
    restoreCanvas();
  }
}, 15_000);

test("a refused seed opens from the state the pages convert, converting once", async () => {
  const restoreCanvas = stubCanvas();
  const conversions = spyOn(conversion, "toProseDoc");
  const refusal = spyOn(
    canonicalSessions,
    "createCanonicalSession",
  ).mockImplementation(() =>
    Result.err(
      new canonicalSessions.CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.dispatch },
        message: "The session cannot hold this document.",
        reason: "refused",
      }),
    ),
  );
  const errors: Error[] = [];
  const mounted = await mountCanonicalEditor("Refused open", errors);
  try {
    await waitForPaintedText(mounted.container, "Refused open");
    expect(refusal).toHaveBeenCalled();
    expect(errors.map(({ message }) => message)).toContain(
      "The session cannot hold this document.",
    );
    expect(mounted.ref().getEditor()?.getView()).toBeNull();
    expect(conversions).toHaveBeenCalledTimes(1);
  } finally {
    await mounted.unmount();
    refusal.mockRestore();
    conversions.mockRestore();
    restoreCanvas();
  }
}, 15_000);
