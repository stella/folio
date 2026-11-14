import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { ensureParaIds } from "@stll/folio-core/docx/ensureParaIds";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { fromMarkdown } from "@stll/folio-core/markdown";

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

const REPLACEMENT_FONT_SIZE = 30;

const docxOf = async (markdown: string, fontSize?: number) => {
  const document = fromMarkdown(markdown);
  if (fontSize !== undefined) {
    const styles = document.package.styles ?? panic("Expected styles");
    styles.docDefaults = {
      ...styles.docDefaults,
      rPr: { ...styles.docDefaults?.rPr, fontSize },
    };
  }
  return (await ensureParaIds(await createDocx(document))).docx;
};

const firstText = (document: ReturnType<DocxEditorRef["getDocument"]>) =>
  JSON.stringify(document?.package.document.content.at(0));

// A handle's methods close over the render that made them, so a host holding
// one across a load read the old document's package around the new body.
test("a handle held across a load reads and saves the document it loaded", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const source = await docxOf("Source first\n\nSource second");
  const replacement = await docxOf(
    "Replacement first\n\nReplacement second",
    REPLACEMENT_FONT_SIZE,
  );
  try {
    await act(async () => {
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <DocxEditor
            experimentalSession="legacy"
            accessibleName="Document editor"
            ref={editor}
            documentBuffer={source}
            showToolbar={false}
          />
        </IntlProvider>,
      );
    });
    await act(async () => editor.current?.ensureEditorView({ focus: false }));
    const held = editor.current ?? panic("Expected an editor handle");

    await act(async () => held.loadDocumentBuffer(replacement));
    await act(async () => held.ensureEditorView({ focus: false }));
    const read = held.getDocument();
    expect(firstText(read)).toContain("Replacement first");
    expect(read?.package.styles?.docDefaults?.rPr?.fontSize).toBe(
      REPLACEMENT_FONT_SIZE,
    );

    // Loading what it read leaves a document that still reads and saves.
    await act(async () =>
      held.loadDocument(read ?? panic("Expected a document")),
    );
    await act(async () => held.ensureEditorView({ focus: false }));
    expect(firstText(held.getDocument())).toContain("Replacement first");
    const saved = await act(async () => held.save());
    const reopened = await parseDocx(saved ?? panic("Expected saved bytes"), {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(firstText(reopened)).toContain("Replacement first");
    expect(reopened.package.styles?.docDefaults?.rPr?.fontSize).toBe(
      REPLACEMENT_FONT_SIZE,
    );
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});
