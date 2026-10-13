import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { parseDocx } from "@stll/folio-core/docx/parser";

import { buildRawPackage, SECTION, WORD_STYLES } from "../../../core/src/__tests__/documentShapes";
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
      `${run("Kept ")}<w:ins ${revision(1)}>${run("inserted")}</w:ins>` +
        `<w:del ${revision(2)}><w:r><w:delText xml:space="preserve"> deleted</w:delText></w:r></w:del>` +
        `${run(" text.")}`,
    ) +
    para(
      `<w:moveFromRangeStart w:id="3" w:name="move" w:author="Reviewer" w:date="${FIXED_DATE}"/>` +
        `<w:moveFrom ${revision(4)}><w:r><w:delText xml:space="preserve">Moved words</w:delText></w:r></w:moveFrom>` +
        `<w:moveFromRangeEnd w:id="3"/> ${run("Here: ")}` +
        `<w:moveToRangeStart w:id="5" w:name="move" w:author="Reviewer" w:date="${FIXED_DATE}"/>` +
        `<w:moveTo ${revision(6)}>${run("Moved words")}</w:moveTo>` +
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
  const modes = [
    { mode: "all-markup", showCommentHighlights: true },
    { mode: "all-markup", showCommentHighlights: false },
    { mode: "final", showCommentHighlights: true },
    { mode: "original", showCommentHighlights: true },
    { mode: "all-markup", showCommentHighlights: true },
  ] as const;
  const visibleByMode: string[] = [];
  const commentIdsByMode: string[][] = [];
  const savedByMode: Uint8Array[] = [];
  let initialDoc: PMNode | null = null;
  try {
    for (const [index, { mode, showCommentHighlights }] of modes.entries()) {
      await act(async () => {
        root.render(
          <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
            <DocxEditor
              ref={editor}
              document={parsedDocument}
              reviewDisplayMode={mode}
              showCommentHighlights={showCommentHighlights}
              readOnly
              pageRenderer="display-list"
              showToolbar={false}
            />
          </IntlProvider>,
        );
      });
      if (index === 0) {
        await act(async () => {
          await editor.current?.ensureEditorView({ focus: false });
          editor.current?.getEditor()?.relayout();
        });
        initialDoc =
          editor.current?.getEditor()?.getView()?.state.doc ??
          panic("Expected the loaded editor document");
      } else {
        expect(editor.current?.getEditor()?.getView()?.state.doc).toBe(initialDoc);
        await act(async () => editor.current?.getEditor()?.relayout());
      }

      const paintedText = [...container.querySelectorAll(".layout-line")]
        .map((line) => line.textContent ?? "")
        .join(" ");
      visibleByMode.push(paintedText);
      commentIdsByMode.push(
        [...container.querySelectorAll(".layout-line[data-comment-id]")].map(
          (line) => line.getAttribute("data-comment-id") ?? "",
        ),
      );
      const saved = await editor.current?.save();
      if (saved === null || saved === undefined) panic("Expected a saved DOCX");
      savedByMode.push(new Uint8Array(saved));
    }

    const allMarkup = visibleByMode.at(0) ?? "";
    const allMarkupWithoutComments = visibleByMode.at(1) ?? "";
    const final = visibleByMode.at(2) ?? "";
    const original = visibleByMode.at(3) ?? "";
    const allMarkupAfterSwitches = visibleByMode.at(4) ?? "";
    expect(allMarkup).toContain("inserted");
    expect(allMarkup).toContain("deleted");
    expect(final).toContain("inserted");
    expect(final).not.toContain("deleted");
    expect(original).not.toContain("inserted");
    expect(original).toContain("deleted");
    expect(allMarkup).toContain("Moved words");
    expect(final).toContain("Moved words");
    expect(original).toContain("Moved words");
    expect(allMarkup).toContain("Inserted row");
    expect(allMarkup).toContain("Deleted row");
    expect(final).toContain("Inserted row");
    expect(final).not.toContain("Deleted row");
    expect(original).not.toContain("Inserted row");
    expect(original).toContain("Deleted row");
    expect(allMarkup).toContain("Comment anchor");
    expect(final).toContain("Numbered clause");
    expect(final).toContain("1.");
    expect(final).not.toContain("§");
    expect(original).toContain("§");
    expect(commentIdsByMode.at(0)).toContain("7");
    expect(commentIdsByMode.at(1)).toEqual([]);
    expect(commentIdsByMode.at(2)).toEqual([]);
    expect(commentIdsByMode.at(3)).toEqual([]);
    expect(commentIdsByMode.at(4)).toContain("7");
    expect(allMarkupWithoutComments).toContain("Comment anchor");
    expect(allMarkupAfterSwitches).toBe(allMarkup);
    const firstSaved = savedByMode.at(0) ?? panic("Expected the initial mode save");
    for (const saved of savedByMode.slice(1)) {
      expect(saved.length).toBe(firstSaved.length);
      expect(saved.every((byte, index) => byte === firstSaved.at(index))).toBe(true);
    }
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});
