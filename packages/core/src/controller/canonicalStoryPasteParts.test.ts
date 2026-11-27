/**
 * Kept markup pasted into a header, footer or note being edited on its own
 * copies the parts it names into that story's own part: a body chart pasted
 * into a header names a fresh relationship of the header part, judged against
 * the header part's relationships rather than the main document's, and the
 * save writes it there. The editor's own copy reads the source package's
 * parts in the background, and a paste made before that read finishes waits
 * for it.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import type { OpStory } from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { buildRawPackage, SECTION } from "../__tests__/documentShapes";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { parseRelationships, resolveRelativePath } from "../docx/relsParser";
import type { CanonicalRefusalGround } from "../types/canonicalCapabilities";
import type { Document } from "../types/document";
import { decodeText } from "../utils/textEncoding";
import { createHeaderFooterEditorManager } from "./headerFooterEditorManager";
import {
  createHiddenEditorManager,
  type HiddenEditorManagerDeps,
} from "./hiddenEditorManager";
import { createNoteEditorManager } from "./noteEditorManager";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";

const CHART_RUN =
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="914400" cy="457200"/><wp:docPr id="31" name="Chart 31"/>' +
  `<a:graphic><a:graphicData uri="${CHART_NS}">` +
  `<c:chart xmlns:c="${CHART_NS}" r:id="rIdChart"/>` +
  "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

const textParagraph = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/**
 * A package whose body holds a chart under the main document's `rIdChart`,
 * with a header, a footer and a footnote. The header's own relationships
 * hold `rIdChart` for something else, so a paste judged against the main
 * document's relationships would name the wrong target there.
 */
const storyPackage = async (): Promise<ArrayBuffer> => {
  const bytes = await buildRawPackage({
    styles: null,
    body:
      textParagraph("Intro") +
      `<w:p><w:r><w:t xml:space="preserve">Chart </w:t></w:r>${CHART_RUN}</w:p>` +
      '<w:p><w:r><w:t xml:space="preserve">Noted</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' +
      SECTION(
        '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>',
      ),
    parts: [
      {
        path: "word/header1.xml",
        contentType: `${WML_CT}.header+xml`,
        relationshipType: "header",
        relationshipId: "rIdHeader1",
        body: `${XML_DECL}<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}">${textParagraph("Header text")}</w:hdr>`,
      },
      {
        path: "word/_rels/header1.xml.rels",
        body: `${XML_DECL}<Relationships xmlns="${REL_NS}"><Relationship Id="rIdChart" Type="${R_NS}/hyperlink" Target="https://header.example.test/" TargetMode="External"/></Relationships>`,
      },
      {
        path: "word/footer1.xml",
        contentType: `${WML_CT}.footer+xml`,
        relationshipType: "footer",
        relationshipId: "rIdFooter1",
        body: `${XML_DECL}<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}">${textParagraph("Footer text")}</w:ftr>`,
      },
      {
        path: "word/footnotes.xml",
        contentType: `${WML_CT}.footnotes+xml`,
        relationshipType: "footnotes",
        relationshipId: "rIdFootnotes",
        body:
          `${XML_DECL}<w:footnotes xmlns:w="${W_NS}" xmlns:r="${R_NS}">` +
          '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
          '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
          '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Footnote text</w:t></w:r></w:p></w:footnote>' +
          "</w:footnotes>",
      },
      {
        path: "word/charts/chart1.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        relationshipType: "chart",
        relationshipId: "rIdChart",
        body: `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:r="${R_NS}"><c:chart/><c:externalData r:id="rId1"/></c:chartSpace>`,
      },
      {
        path: "word/charts/_rels/chart1.xml.rels",
        body: `${XML_DECL}<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${R_NS}/package" Target="../embeddings/workbook1.xlsx"/></Relationships>`,
      },
      {
        path: "word/embeddings/workbook1.xlsx",
        contentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        body: "PK workbook bytes",
      },
    ],
  });
  return bytes.slice().buffer;
};

type StoryCase = {
  name: string;
  story: OpStory;
  /** The key the story part's relationships are kept under. */
  key: string;
  part: string;
  text: string;
};

const CASES: readonly StoryCase[] = [
  {
    name: "header",
    story: { kind: "header", rId: "rIdHeader1" },
    key: "header:rIdHeader1",
    part: "word/header1.xml",
    text: "Header text",
  },
  {
    name: "footer",
    story: { kind: "footer", rId: "rIdFooter1" },
    key: "footer:rIdFooter1",
    part: "word/footer1.xml",
    text: "Footer text",
  },
  {
    name: "footnote",
    story: { kind: "footnote", id: 1 },
    key: "footnotes",
    part: "word/footnotes.xml",
    text: "Footnote text",
  },
];

/** The position just after `needle` in the first text holding it. */
const positionAfter = (doc: PMNode, needle: string): number => {
  let found: number | undefined;
  doc.descendants((node, position) => {
    if (found !== undefined) return false;
    const index = node.isText ? (node.text ?? "").indexOf(needle) : -1;
    if (index >= 0) found = position + index + needle.length;
    return found === undefined;
  });
  return found ?? panic(`The story holds no "${needle}".`);
};

/** The content of the paragraph holding `needle`, as a document range. */
const paragraphHolding = (doc: PMNode, needle: string) => {
  let range: { from: number; to: number } | undefined;
  doc.descendants((node, position) => {
    if (range !== undefined) return false;
    if (node.type.name !== "paragraph") return true;
    if (node.textContent.includes(needle))
      range = { from: position + 1, to: position + node.nodeSize - 1 };
    return false;
  });
  return range ?? panic(`No paragraph holds "${needle}".`);
};

const mount = async (story: OpStory, suggesting: boolean) => {
  const releaseDom = acquireHarnessDom();
  const bodyHost = document.createElement("div");
  const storyHost = document.createElement("div");
  document.body.append(bodyHost, storyHost);
  const source: Document = await parseDocx(await storyPackage(), {
    preloadFonts: false,
    detectVariables: false,
  });
  const refused: { message: string; ground: CanonicalRefusalGround }[] = [];
  const onSessionRefusal = (message: string, ground: CanonicalRefusalGround) =>
    refused.push({ message, ground });
  const deps = {
    getHost: () => bodyHost,
    getDocument: () => source,
    getStyles: () => source.package.styles ?? null,
    getEmbeddedFontFamilyMap: () => undefined,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getReadOnly: () => false,
    getDocumentIdentity: () => "story-paste-parts",
    getDocumentContext: () => source,
    getEditingMode: () => (suggesting ? "suggesting" : "editing"),
    getSuggestionAuthor: () => "Paster",
    onTransaction: () => undefined,
    onCanonicalCompositionLayout: () => undefined,
    onSelectionChange: () => undefined,
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => undefined,
    onEditorViewReady: () => undefined,
    onEditorViewDestroy: () => undefined,
    onRemoteSelectionsChange: () => undefined,
    onSessionRefusal,
  } satisfies HiddenEditorManagerDeps;
  const manager = createHiddenEditorManager(deps);
  const storyDeps = {
    getHost: () => storyHost,
    getDocument: () => source,
    getCanonicalApi: () => manager.api,
    readCanonicalDocument: () => manager.readCanonicalDocument(),
    getStyles: () => source.package.styles ?? null,
    getTheme: () => null,
    onSessionRefusal,
  };
  const parts = createHeaderFooterEditorManager(storyDeps);
  const notes = createNoteEditorManager(storyDeps);
  manager.ensureView();
  parts.sync();
  const body = manager.getView() ?? panic("The body editor did not mount.");
  let storyView: EditorView | null = null;
  if (story !== "main" && (story.kind === "header" || story.kind === "footer"))
    storyView = parts.getView(story.rId);
  if (
    story !== "main" &&
    (story.kind === "footnote" || story.kind === "endnote")
  )
    storyView = notes.activate({ kind: story.kind, noteId: story.id });
  const view = storyView ?? panic("The story editor did not mount.");
  const dispose = () => {
    parts.destroy();
    notes.destroy();
    manager.destroyView();
    bodyHost.remove();
    storyHost.remove();
    releaseDom();
  };
  return { manager, body, view, refused, dispose };
};

/** Copy the body's chart paragraph through the editor's own copy. */
const copyChart = (body: EditorView) => {
  const { from, to } = paragraphHolding(body.state.doc, "Chart");
  body.dispatch(
    body.state.tr.setSelection(TextSelection.create(body.state.doc, from, to)),
  );
  const clipboardData = new DataTransfer();
  body.dom.dispatchEvent(
    new ClipboardEvent("copy", {
      bubbles: true,
      cancelable: true,
      clipboardData,
    }),
  );
  return { clipboardData, slice: body.state.selection.content() };
};

/** Paste into the story editor after its text, as the browser hands it over. */
const pasteInto = (
  view: EditorView,
  text: string,
  { clipboardData, slice }: ReturnType<typeof copyChart>,
) => {
  const at = positionAfter(view.state.doc, text);
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, at)),
  );
  return view.someProp("handlePaste", (handle) =>
    handle(
      view,
      new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData,
      }),
      slice,
    ),
  );
};

/** Wait for the background read of the copy's parts and the paste after it. */
const settle = async (done: () => boolean) => {
  for (let attempt = 0; attempt < 200 && !done(); attempt += 1)
    // oxlint-disable-next-line no-await-in-loop -- polls the paste's completion
    await new Promise((resolve) => setTimeout(resolve, 5));
};

const save = async (manager: ReturnType<typeof createHiddenEditorManager>) => {
  const snapshot =
    manager.api.captureCanonicalSave() ?? panic("The session has no save.");
  const saved = await serializeCanonicalSave({
    snapshot,
    options: { mode: "full" },
  });
  return saved.buffer;
};

const textOf = async (zip: ZipArchive, path: string) =>
  decodeText((await zip.entry(path)?.bytes()) ?? new Uint8Array(), {
    ignoreBOM: true,
  });

/** The relationships part of the part at `path`. */
const relsPathOf = (path: string) => {
  const slash = path.lastIndexOf("/");
  return `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
};

/** Every internal relationship of every part resolves to a part. */
const assertResolves = async (zip: ZipArchive) => {
  const paths = new Set(zip.paths());
  for (const relsPath of zip.paths().filter((path) => path.endsWith(".rels")))
    for (const relationship of parseRelationships(
      // oxlint-disable-next-line no-await-in-loop -- each part is read once
      await textOf(zip, relsPath),
    ).values())
      if (relationship.targetMode !== "External")
        expect(
          paths.has(resolveRelativePath(relsPath, relationship.target)),
          `${relsPath} ${relationship.id}`,
        ).toBe(true);
};

/** The ids the chart markup in `xml` names. */
const chartIds = (xml: string) =>
  [...xml.matchAll(/<c:chart [^>]*r:id="(?<id>[^"]+)"/gu)].flatMap(
    ({ groups }) => (groups?.["id"] === undefined ? [] : [groups["id"]]),
  );

describe.each(["editing", "suggesting"] as const)("in %s", (mode) => {
  for (const { name, story, key, part, text } of CASES)
    test(`a body chart pasted into the ${name} before its copy's parts are read copies them into the ${name} part's relationships, resolves after a reopen, and undo removes the copies`, async () => {
      const editor = await mount(story, mode === "suggesting");
      try {
        const before = editor.manager.api.getCanonicalDocument();
        const shown = editor.view.state.doc;
        const copied = copyChart(editor.body);
        // The copy's parts are still being read: the paste waits for them.
        expect(pasteInto(editor.view, text, copied)).toBe(true);
        expect(editor.manager.api.getCanonicalDocument()).toEqual(before);
        expect(editor.view.state.doc.eq(shown)).toBe(true);
        await settle(
          () => editor.refused.length > 0 || !editor.view.state.doc.eq(shown),
        );
        expect(editor.refused).toEqual([]);
        const pasted =
          editor.manager.api.getCanonicalDocument() ??
          panic("The session has no document.");
        const held =
          pasted.package.annotationRelationships?.get(key) ??
          panic(`The ${name} part has no relationships.`);
        const copies = [...held.values()].filter(({ type }) =>
          type.endsWith("/chart"),
        );
        expect(copies).toHaveLength(1);
        const copy = copies.at(0) ?? panic("The chart was not copied.");
        expect(copy.id).not.toBe("rIdChart");
        expect(copy.target).toBe("charts/chart2.xml");

        const zip = await ZipArchive.read(await save(editor.manager));
        await assertResolves(zip);
        // The story names the copy through its own part's relationships.
        const partXml = await textOf(zip, part);
        expect(chartIds(partXml)).toEqual([copy.id]);
        const partRelationships = parseRelationships(
          await textOf(zip, relsPathOf(part)),
        );
        expect(partRelationships.get(copy.id)?.target).toBe(
          "charts/chart2.xml",
        );
        if (name === "header")
          expect(partRelationships.get("rIdChart")?.target).toBe(
            "https://header.example.test/",
          );
        expect(
          parseRelationships(
            await textOf(zip, "word/charts/_rels/chart2.xml.rels"),
          ).get("rId1")?.target,
        ).toBe("../embeddings/workbook2.xlsx");
        // The body's chart keeps its own.
        expect(chartIds(await textOf(zip, "word/document.xml"))).toEqual([
          "rIdChart",
        ]);
        const reopened = await parseDocx(
          (await zip.generate({ compression: "store" })).buffer,
          { preloadFonts: false, detectVariables: false },
        );
        expect(
          reopened.package.annotationRelationships?.get(key)?.get(copy.id)
            ?.target,
        ).toBe("charts/chart2.xml");

        expect(
          editor.manager.api.applyCanonicalStoryHistory({
            view: editor.view,
            story,
            direction: "undo",
          }),
        ).toBe(true);
        expect(editor.manager.api.getCanonicalDocument()).toEqual(before);
        const undone = await ZipArchive.read(await save(editor.manager));
        expect(undone.paths()).not.toContain("word/charts/chart2.xml");
        expect(chartIds(await textOf(undone, part))).toEqual([]);
      } finally {
        editor.dispose();
      }
    });
});

test("a paste waiting on its copy's parts refuses as stale once the story editor moved on", async () => {
  const editor = await mount({ kind: "header", rId: "rIdHeader1" }, false);
  try {
    const before = editor.manager.api.getCanonicalDocument();
    expect(pasteInto(editor.view, "Header", copyChart(editor.body))).toBe(true);
    // The selection moves before the read finishes.
    editor.view.dispatch(
      editor.view.state.tr.setSelection(
        TextSelection.create(editor.view.state.doc, 1),
      ),
    );
    await settle(() => editor.refused.length > 0);
    expect(editor.refused.map(({ ground }) => ground)).toEqual([
      { condition: "stale" },
    ]);
    expect(editor.manager.api.getCanonicalDocument()).toEqual(before);
  } finally {
    editor.dispose();
  }
});
