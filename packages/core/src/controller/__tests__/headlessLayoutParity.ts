import { afterEach, beforeEach, expect, setSystemTime } from "bun:test";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";

import { runLayoutPipeline, type LayoutPipelineDeps } from "../layoutPipeline";
import { createLayoutSession } from "../layoutSession";
import { parseDocx } from "../../docx/parser";
import { layoutDocxHeadless } from "../../headless-layout";
import { convertHeaderFooterToContent } from "../../layout-bridge/convert/headerFooterLayout";
import { getColumns } from "../../layout-bridge/sectionColumns";
import { resolveSectionHeaderFooterRefs } from "../../layout-engine/headerFooterRefs";
import { clearAllCaches } from "../../layout-engine/measure/cache";
import {
  installCanvasMeasureProvider,
  resetCanvasContext,
} from "../../layout-engine/measure/measureContainer";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "../../layout-engine/measure/measureProvider";
import type { HeaderFooterContent, Layout } from "../../layout-engine/types";
import { LayoutSelectionGate } from "../../paged-layout/LayoutSelectionGate";
import { getMargins, getPageSize } from "../../paged-layout/sectionGeometry";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { resolveHeaderFooterContent } from "../../utils/headerFooter";

/** Keep the same deterministic backend alive across both asynchronous paths. */
export const installParityMeasurement = (): void => {
  let originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  let originalProvider = getMeasureProvider();
  beforeEach(() => {
    setSystemTime(new Date("2026-01-01T00:00:00Z"));
    originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
    originalProvider = getMeasureProvider();
    Object.defineProperty(globalThis, "document", {
      configurable: true,
      value: {
        createElement: () => ({
          getContext: () => ({
            font: "",
            measureText: (text: string) => ({
              width: [...text].length * 7,
              actualBoundingBoxAscent: 8,
              actualBoundingBoxDescent: 2,
            }),
          }),
        }),
      },
    });
    clearAllCaches();
    resetCanvasContext();
    installCanvasMeasureProvider();
  });
  afterEach(() => {
    setSystemTime();
    setMeasureProvider(originalProvider);
    resetCanvasContext();
    clearAllCaches();
    if (originalDocument) {
      Object.defineProperty(globalThis, "document", originalDocument);
    } else {
      Reflect.deleteProperty(globalThis, "document");
    }
  });
};

export const pageBreaks = (layout: Layout) =>
  layout.pages.map(({ fragments }) => ({
    first: fragments.at(0)?.blockId,
    last: fragments.at(-1)?.blockId,
  }));

/** Adapter wiring with no painter; conversion, measurement and pagination are real. */
export const layoutBothPaths = async (bytes: ArrayBuffer) => {
  const headlessResult = await layoutDocxHeadless(bytes);
  if (headlessResult.isErr()) {
    return panic("Headless fixture failed to load or lay out", { cause: headlessResult.error });
  }
  // Parse independently, as an editor load does, rather than reusing the
  // headless result's projection or prepared furniture in the editor oracle.
  const document = await parseDocx(bytes, { preloadFonts: false });
  const sectionProperties =
    document.package.document.sections?.at(0)?.properties ??
    document.package.document.finalSectionProperties;
  const margins = getMargins(sectionProperties);
  const pageSize = getPageSize(sectionProperties);
  const hf = resolveHeaderFooterContent(document.package);
  const deps = {
    contentWidth: pageSize.w - margins.left - margins.right,
    columns: getColumns(sectionProperties),
    pageSize,
    margins,
    pageGap: 0,
    showMarginGuides: false,
    marginGuideColor: undefined,
    syncCoordinator: new LayoutSelectionGate(),
    headerContent: hf.headerContent,
    footerContent: hf.footerContent,
    firstPageHeaderContent: hf.firstPageHeaderContent,
    firstPageFooterContent: hf.firstPageFooterContent,
    headerContentRId: hf.activeHeaderRId,
    footerContentRId: hf.activeFooterRId,
    firstPageHeaderContentRId: hf.activeFirstHeaderRId,
    firstPageFooterContentRId: hf.activeFirstFooterRId,
    sectionHeaderFooterRefs: resolveSectionHeaderFooterRefs(document),
    theme: document.package.theme,
    sectionProperties,
    document,
    defaultTabStop: document.package.settings?.defaultTabStop,
    mirrorMargins: document.package.settings?.mirrorMargins === true,
    styles: document.package.styles,
    layout: null,
    hfPMs: null,
    painter: null,
    pagesContainer: null,
    session: createLayoutSession(),
    renderHfFromContentOrPm: (part, rId, _pms, width, metrics, options) =>
      convertHeaderFooterToContent(part, width, metrics, {
        ...options,
        ...(rId ? { rId } : {}),
      }),
    renderHeaderFooterContentByRId: (parts, _pms, width, metrics, options) => {
      const contents = new Map<string, HeaderFooterContent>();
      for (const [rId, part] of parts ?? []) {
        const content = convertHeaderFooterToContent(part, width, metrics, { ...options, rId });
        if (content) contents.set(rId, content);
      }
      return contents;
    },
    readFontSetSignature: () => "layout-parity-fixed-fonts",
    buildFootnoteRenderItems: () => new Map(),
    describeInvalidHighlightMarks: () => "",
    emptyTemplatePreviewEntries: [],
    emptyTemplatePreviewHidden: [],
    hyphenationReadiness: { track: () => undefined, cancel: () => undefined },
    markupView: "all-markup",
  } satisfies LayoutPipelineDeps<null>;
  const editor = runLayoutPipeline(
    deps,
    EditorState.create({
      doc: toProseDoc(document, {
        ...(document.package.styles ? { styles: document.package.styles } : {}),
        ...(document.package.theme ? { theme: document.package.theme } : {}),
      }),
    }),
  );
  if (!editor.layout) return panic("Editor fixture failed to lay out");
  const headless = headlessResult.value;
  expect(headless.layout.pages.length).toBeGreaterThan(0);
  expect(editor.layout.pages).toHaveLength(headless.layout.pages.length);
  expect(pageBreaks(editor.layout)).toEqual(pageBreaks(headless.layout));
  return { headless, editor: editor.layout };
};
