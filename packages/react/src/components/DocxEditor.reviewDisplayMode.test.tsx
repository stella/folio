import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import type { EditorReadOnlyReason } from "@stll/folio-core/managers/EditorModeManager";
import type { Node as PMNode } from "prosemirror-model";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { clearAllCaches } from "@stll/folio-core/layout-engine/measure/cache";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "@stll/folio-core/layout-engine/measure/measureProvider";
import { resetCanvasContext } from "@stll/folio-core/layout-engine/measure/measureContainer";
import { sizeProportionalCharWidth } from "@stll/folio-core/layout-engine/measure/__tests__/fakeTextMeasure";

import { buildRawPackage, SECTION, WORD_STYLES } from "@stll/folio-core/__tests__/documentShapes";
import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const FIXED_DATE = "2026-01-01T00:00:00Z";
const revision = (id: number): string => `w:id="${id}" w:author="Reviewer" w:date="${FIXED_DATE}"`;
const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (content: string, properties = ""): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}${content}</w:p>`;

const reviewDocument = async (): Promise<Uint8Array> => {
  const body =
    para(
      `${run("Kept ")}<w:ins ${revision(1)}>${run("TRACKED_INSERTION")}</w:ins>` +
        `<w:del ${revision(2)}><w:r><w:delText xml:space="preserve">TRACKED_DELETION</w:delText></w:r></w:del>` +
        `${run(" text.")}`,
    ) +
    para(
      `<w:moveFromRangeStart w:id="3" w:name="move" w:author="Reviewer" w:date="${FIXED_DATE}"/>` +
        `<w:moveFrom ${revision(4)}><w:r><w:delText xml:space="preserve">MOVE_FROM</w:delText></w:r></w:moveFrom>` +
        `<w:moveFromRangeEnd w:id="3"/> ${run("Here: ")}` +
        `<w:moveToRangeStart w:id="5" w:name="move" w:author="Reviewer" w:date="${FIXED_DATE}"/>` +
        `<w:moveTo ${revision(6)}>${run("MOVE_TO")}</w:moveTo>` +
        `<w:moveToRangeEnd w:id="5"/>`,
    ) +
    para(
      `<w:commentRangeStart w:id="7"/>${run("Comment anchor")}` +
        `<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>`,
    ) +
    para(
      `<w:r><w:rPr><w:sz w:val="32"/><w:rPrChange ${revision(8)}><w:rPr/></w:rPrChange></w:rPr><w:t>Tracked format</w:t></w:r>`,
    ) +
    para(
      run("Numbered clause"),
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:pPrChange ${revision(9)}><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr></w:pPrChange>`,
    ) +
    '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' +
    `<w:tr><w:trPr><w:ins ${revision(10)}/></w:trPr><w:tc><w:p><w:ins ${revision(11)}>${run("Inserted row")}</w:ins></w:p></w:tc><w:tc><w:p>${run("Cell B")}</w:p></w:tc></w:tr>` +
    `<w:tr><w:trPr><w:del ${revision(12)}/></w:trPr><w:tc><w:p><w:del ${revision(13)}><w:r><w:delText>Deleted row</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:p>${run("Cell D")}</w:p></w:tc></w:tr></w:tbl>` +
    SECTION();
  return buildRawPackage({
    body,
    styles: WORD_STYLES,
    parts: [
      {
        path: "word/comments.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        relationshipType: "comments",
        relationshipId: "rIdComments",
        body:
          `<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
          `<w:comment w:id="7" w:author="Reviewer" w:date="${FIXED_DATE}"><w:p>${run("Please review this clause.")}</w:p></w:comment></w:comments>`,
      },
      {
        path: "word/numbering.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        relationshipType: "numbering",
        relationshipId: "rIdNumbering",
        body:
          '<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
          '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="§"/></w:lvl></w:abstractNum>' +
          '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>' +
          '<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>',
      },
    ],
  });
};

const createEditAttemptRecorder =
  (reasons: EditorReadOnlyReason[]) => (reason: EditorReadOnlyReason) => {
    reasons.push(reason);
  };

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

test("host review display modes change visible layout without changing saved DOCX bytes", async () => {
  const bytes = await reviewDocument();
  const parsedDocument = await parseDocx(bytes, { preloadFonts: false, detectVariables: false });
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const refusals: EditorReadOnlyReason[] = [];
  const onReadonlyEditAttempt = createEditAttemptRecorder(refusals);
  const modes = [
    { mode: "all-markup", showCommentHighlights: true },
    { mode: "all-markup", showCommentHighlights: false },
    { mode: "final", showCommentHighlights: true },
    { mode: "original", showCommentHighlights: true },
    { mode: "all-markup", showCommentHighlights: true },
  ] as const;
  const visibleByMode: string[] = [];
  const tableRowCountByMode: number[] = [];
  const commentIdsByMode: string[][] = [];
  const savedByMode: Uint8Array[] = [];
  let initialDoc: PMNode | null = null;
  const previousMeasureProvider = getMeasureProvider();
  const originalGetContext =
    Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, "getContext") ??
    panic("Expected the canvas context method");
  const canvasContext = {
    font: "",
    fontKerning: "auto",
    measureText(text: string) {
      let width = 0;
      for (const char of text) width += sizeProportionalCharWidth(char, this.font);
      return { width, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 };
    },
  };
  const relayout = () => {
    const controller = editor.current?.getEditor() ?? panic("Expected the editor controller");
    let completed = false;
    const unsubscribe = controller.on("layoutComplete", () => {
      completed = true;
    });
    controller.relayout();
    unsubscribe();
    expect(completed).toBe(true);
  };
  try {
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => canvasContext,
    });
    resetCanvasContext();
    clearAllCaches();
    for (const [index, { mode, showCommentHighlights }] of modes.entries()) {
      await act(async () => {
        root.render(
          <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
            <DocxEditor
              ref={editor}
              document={parsedDocument}
              reviewDisplayMode={mode}
              showCommentHighlights={showCommentHighlights}
              pageRenderer="display-list"
              showToolbar={false}
              onReadonlyEditAttempt={onReadonlyEditAttempt}
            />
          </IntlProvider>,
        );
      });
      if (index === 0) {
        await act(async () => {
          editor.current?.ensureEditorView({ focus: false });
        });
        await act(async () => relayout());
        initialDoc =
          editor.current?.getEditor()?.getView()?.state.doc ??
          panic("Expected the loaded editor document");
      } else {
        expect(editor.current?.getEditor()?.getView()?.state.doc).toBe(initialDoc);
        await act(async () => relayout());
      }

      const controller = editor.current?.getEditor() ?? panic("Expected the editor controller");
      const view = controller.getView() ?? panic("Expected the body editor view");
      if (mode === "final" || mode === "original") {
        expect(editor.current?.getReadOnlyReason()).toBe("review-display-mode");
        expect(view.editable).toBe(false);
        const authored = view.state.doc;
        const event = new InputEvent("beforeinput", {
          inputType: "insertText",
          data: "refused",
          bubbles: true,
          cancelable: true,
        });
        await act(async () => {
          view.dom.dispatchEvent(event);
        });
        expect(event.defaultPrevented).toBe(true);
        expect(refusals.at(-1)).toBe("review-display-mode");
        expect(view.state.doc).toBe(authored);
      } else {
        expect(editor.current?.getReadOnlyReason()).toBeNull();
        expect(view.editable).toBe(true);
      }
      const layout = controller.getLayout() ?? panic("Expected the settled page layout");
      visibleByMode.push(
        layout.pages
          .map(
            (_, pageIndex) =>
              editor.current?.getPageText(pageIndex + 1) ??
              panic("Expected page text for each laid-out page"),
          )
          .join("\n"),
      );
      tableRowCountByMode.push(
        layout.pages
          .flatMap((page) => page.fragments)
          .filter((fragment) => fragment.kind === "table")
          .reduce((total, fragment) => total + fragment.toRow - fragment.fromRow, 0),
      );
      commentIdsByMode.push(
        [...container.querySelectorAll(".layout-line[data-comment-id]")].map(
          (line) => line.getAttribute("data-comment-id") ?? "",
        ),
      );
      let saved: ArrayBuffer | null | undefined;
      await act(async () => {
        saved = await editor.current?.save();
      });
      if (saved === null || saved === undefined) panic("Expected a saved DOCX");
      savedByMode.push(new Uint8Array(saved));
    }

    const allMarkup = visibleByMode.at(0) ?? "";
    const allMarkupWithoutComments = visibleByMode.at(1) ?? "";
    const final = visibleByMode.at(2) ?? "";
    const original = visibleByMode.at(3) ?? "";
    const allMarkupAfterSwitches = visibleByMode.at(4) ?? "";
    expect(allMarkup).toContain("TRACKED_INSERTION");
    expect(allMarkup).toContain("TRACKED_DELETION");
    expect(final).toContain("TRACKED_INSERTION");
    expect(final).not.toContain("TRACKED_DELETION");
    expect(original).not.toContain("TRACKED_INSERTION");
    expect(original).toContain("TRACKED_DELETION");
    expect(allMarkup).toContain("MOVE_FROM");
    expect(allMarkup).toContain("MOVE_TO");
    expect(final).not.toContain("MOVE_FROM");
    expect(final).toContain("MOVE_TO");
    expect(original).toContain("MOVE_FROM");
    expect(original).not.toContain("MOVE_TO");
    expect(allMarkup).toContain("Comment anchor");
    expect(final).toContain("Numbered clause");
    expect(tableRowCountByMode.at(0)).toBe(2);
    expect(tableRowCountByMode.at(2)).toBe(1);
    expect(tableRowCountByMode.at(3)).toBe(1);
    expect(commentIdsByMode.at(0)).toContain("7");
    expect(commentIdsByMode.at(1)).toEqual([]);
    expect(commentIdsByMode.at(2)).toEqual([]);
    expect(commentIdsByMode.at(3)).toEqual([]);
    expect(commentIdsByMode.at(4)).toContain("7");
    expect(allMarkupWithoutComments).toBe(allMarkup);
    expect(allMarkupAfterSwitches).toBe(allMarkup);
    const firstSaved = savedByMode.at(0) ?? panic("Expected the initial mode save");
    for (const saved of savedByMode.slice(1)) {
      expect(saved.length).toBe(firstSaved.length);
      expect(saved.every((byte, index) => byte === firstSaved.at(index))).toBe(true);
    }
  } finally {
    await act(async () => root.unmount());
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", originalGetContext);
    setMeasureProvider(previousMeasureProvider);
    resetCanvasContext();
    clearAllCaches();
    container.remove();
  }
});
