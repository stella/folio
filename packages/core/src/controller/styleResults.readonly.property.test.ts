/**
 * The style resolver hands out the document's own style values and, through
 * the engine, results it caches and shares between callers. This file runs
 * the resolver and the engine over deep-frozen style packages, and the caller
 * paths the hosts use over an engine whose every result is deep-frozen too:
 * none may throw, and each must read what it reads from a plain engine over
 * the same package.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { deepFreeze } from "../__tests__/deepFreeze";
import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { getCachedNumberingMap } from "../docx/numberingParser";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { parseDocx } from "../docx/parser";
import { resolveParagraphMarkFormatting } from "../layout-bridge/convert/paragraphMarkFormatting";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import {
  paragraphRunStyleContext,
  resolveEffectiveRunStyleFormatting,
  resolveParagraphInheritedRunFormatting,
} from "../prosemirror/runStyleFormatting";
import { singletonManager } from "../prosemirror/schema";
import {
  paragraphStyleCascadeAttrs,
  resolveParagraphDefaultTextFormatting,
  tableCellParagraphOverlay,
} from "../prosemirror/styles/paragraphStyleCascade";
import {
  listAttrsFromResolvedStyle,
  paragraphAttrsFromResolvedStyle,
} from "../prosemirror/styles/resolvedStyleAttrs";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import { createStyleEngine, type StyleEngine } from "../style-engine";
import type { Document } from "../types/document";
import type { ReadonlyDocument } from "../types/documentModel";
import { createHiddenEditorState } from "./hiddenEditorManager";

setDefaultTimeout(propertyTestTimeout(120_000));

const loadShapes = () =>
  Promise.all(
    DOCUMENT_SHAPES.map(async (shape) => ({
      id: shape.id,
      document: await parseDocx(
        (
          await prepareCanonicalDocxInput((await shape.build()).slice().buffer)
        ).unwrap(),
        { preloadFonts: false, detectVariables: false },
      ),
    })),
  );

/**
 * An engine whose every result is deep-frozen before it is handed out, so a
 * caller that writes into a result, a cached one included, throws at the write.
 */
const frozenResults = (engine: StyleEngine): StyleEngine => ({
  getStyle: (styleId) => deepFreeze(engine.getStyle(styleId)),
  hasStyle: (styleId) => engine.hasStyle(styleId),
  getDocDefaults: () => deepFreeze(engine.getDocDefaults()),
  getDefaultParagraphStyle: () => deepFreeze(engine.getDefaultParagraphStyle()),
  getDefaultCharacterStyle: () => deepFreeze(engine.getDefaultCharacterStyle()),
  getDefaultTableStyle: () => deepFreeze(engine.getDefaultTableStyle()),
  getParagraphStyles: () => deepFreeze(engine.getParagraphStyles()),
  getTableStyles: () => deepFreeze(engine.getTableStyles()),
  get builtInStyles() {
    return engine.builtInStyles;
  },
  resolveParagraphStyle: (styleId) =>
    deepFreeze(engine.resolveParagraphStyle(styleId)),
  resolveParagraphStyleInTable: (styleId, overlay) =>
    deepFreeze(engine.resolveParagraphStyleInTable(styleId, overlay)),
  resolveRunStyle: (styleId) => deepFreeze(engine.resolveRunStyle(styleId)),
  getRunStyleOwnProperties: (styleId) =>
    deepFreeze(engine.getRunStyleOwnProperties(styleId)),
  invalidate: () => engine.invalidate(),
  stats: () => engine.stats(),
});

const TABLE_REGIONS = [
  [],
  ["firstRow"],
  ["band1Horz", "firstCol"],
  ["lastRow", "lastCol", "seCell"],
] as const;

const styleIdsOf = (document: ReadonlyDocument): (string | null)[] => [
  null,
  "missing-style",
  ...(document.package.styles?.styles.map(({ styleId }) => styleId) ?? []),
];

/** Every resolver and engine read of the package's styles. */
const readResolver = (document: ReadonlyDocument, cache: boolean) => {
  const styles = document.package.styles;
  const resolver = createStyleResolver(styles);
  const engine = createStyleEngine(styles, { cache });
  return styleIdsOf(document).map((styleId) => ({
    style: styleId === null ? undefined : resolver.getStyle(styleId),
    paragraph: resolver.resolveParagraphStyle(styleId),
    inTable: resolver.resolveParagraphStyleInTable(styleId, {
      spaceAfter: 0,
    }),
    run: resolver.resolveRunStyle(styleId),
    own: resolver.getRunStyleOwnProperties(styleId),
    next: resolver.getNextStyleId(styleId),
    engineParagraph: engine.resolveParagraphStyle(styleId),
    engineRun: engine.resolveRunStyle(styleId),
    engineOwn: engine.getRunStyleOwnProperties(styleId),
    paragraphStyles: engine.getParagraphStyles(),
    tableStyles: engine.getTableStyles(),
    docDefaults: engine.getDocDefaults(),
    defaults: [
      engine.getDefaultParagraphStyle(),
      engine.getDefaultCharacterStyle(),
      engine.getDefaultTableStyle(),
    ],
    builtIn: engine.builtInStyles.styleIdForHeadingLevel(1),
  }));
};

const paragraphNodes = (doc: PMNode): PMNode[] => {
  const paragraphs: PMNode[] = [];
  doc.descendants((node) => {
    if (node.type.name === "paragraph") paragraphs.push(node);
  });
  return paragraphs;
};

/** The caller paths the hosts reach the engine through, with `engine`. */
const readCallers = (document: ReadonlyDocument, engine: StyleEngine) => {
  const styles = document.package.styles;
  const numbering = document.package.numbering;
  const numberingMap = numbering ? getCachedNumberingMap(numbering) : null;
  const state = createHiddenEditorState({
    document: cloneDocumentWithParagraphPropertySources(document),
    styles: styles ?? null,
    manager: singletonManager,
  });
  const paragraphs = paragraphNodes(state.doc);
  return {
    styleAttrs: styleIdsOf(document).map((styleId) => {
      const resolved = engine.resolveParagraphStyle(styleId);
      return {
        attrs: paragraphAttrsFromResolvedStyle(resolved, {
          styleId: styleId ?? "Normal",
        }),
        list: listAttrsFromResolvedStyle(resolved, numberingMap),
        inherited: resolveParagraphInheritedRunFormatting({
          styleId: styleId ?? undefined,
          styleResolver: engine,
          tableRunFormatting: { italic: true },
          paragraphMarkFormatting: { bold: true },
        }).inheritedFormatting,
        defaults: resolveParagraphDefaultTextFormatting(
          styleId ?? undefined,
          { runProperties: { underline: { style: "single" } } },
          engine,
        ),
        tableOverlays: TABLE_REGIONS.map((regions) =>
          tableCellParagraphOverlay(engine, styleId, regions),
        ),
      };
    }),
    paragraphs: paragraphs.map((paragraph) => {
      const attrs = expectParagraphAttrs(paragraph);
      const context = paragraphRunStyleContext(paragraph, engine);
      const runs: unknown[] = [];
      paragraph.descendants((child) => {
        if (child.isText) {
          runs.push(
            resolveEffectiveRunStyleFormatting({
              marks: child.marks,
              paragraphFormatting: context.paragraphFormatting,
              styleResolver: engine,
            }),
          );
        }
      });
      return {
        cascade: paragraphStyleCascadeAttrs({
          styleId: attrs.styleId ?? undefined,
          formatting: attrs._originalFormatting,
          styleResolver: engine,
          includeParagraphMarkRunProperties: true,
        }),
        context,
        mark: resolveParagraphMarkFormatting(attrs, engine),
        runs,
      };
    }),
  };
};

/** What a host's hidden editor and layout read from the model's styles. */
const readHostPaths = (document: ReadonlyDocument) => {
  const styles = document.package.styles;
  const state = createHiddenEditorState({
    document: cloneDocumentWithParagraphPropertySources(document),
    styles: styles ?? null,
    manager: singletonManager,
  });
  return {
    hiddenEditor: state.doc.toJSON(),
    flow: toFlowBlocks(state.doc, styles ? { styles } : {}),
  };
};

const frozenCopy = (document: Document): ReadonlyDocument =>
  deepFreeze(cloneDocumentWithParagraphPropertySources(document));

describe("style resolver results are read-only", () => {
  test("a frozen package refuses a write through a resolver result", async () => {
    const shapes = await loadShapes();
    const styled = shapes.find(({ document }) =>
      document.package.styles?.styles.some(({ rPr }) => rPr !== undefined),
    );
    if (!styled) throw new Error("No shape has a run style.");
    const frozen = frozenCopy(styled.document);
    const withRun = frozen.package.styles?.styles.find(
      ({ rPr }) => rPr !== undefined,
    );
    const own = createStyleResolver(
      frozen.package.styles,
    ).getRunStyleOwnProperties(withRun?.styleId);
    expect(own).toBeDefined();
    expect(() => Object.assign(own ?? {}, { bold: true })).toThrow(TypeError);
  });

  test("the resolver and the engine read a frozen package as they read a plain one", async () => {
    const shapes = await loadShapes();
    await assertProperty(
      fc.property(
        fc.constantFrom(...shapes),
        fc.boolean(),
        ({ document }, cache) => {
          expect(readResolver(frozenCopy(document), cache)).toEqual(
            readResolver(
              cloneDocumentWithParagraphPropertySources(document),
              cache,
            ),
          );
        },
      ),
      { id: "style resolver never writes the package" },
    );
  });

  test("every caller path reads frozen results as it reads plain ones", async () => {
    const shapes = await loadShapes();
    await assertProperty(
      fc.property(fc.constantFrom(...shapes), ({ document }) => {
        const frozen = frozenCopy(document);
        const plain = cloneDocumentWithParagraphPropertySources(document);
        expect(
          readCallers(
            frozen,
            frozenResults(createStyleEngine(frozen.package.styles)),
          ),
        ).toEqual(readCallers(plain, createStyleEngine(plain.package.styles)));
      }),
      { id: "style result callers never write a result" },
    );
  });

  test("the hidden editor and the layout read a frozen model's styles as they read plain ones", async () => {
    const shapes = await loadShapes();
    await assertProperty(
      fc.property(fc.constantFrom(...shapes), ({ document }) => {
        expect(readHostPaths(frozenCopy(document))).toEqual(
          readHostPaths(cloneDocumentWithParagraphPropertySources(document)),
        );
      }),
      { id: "host style readers never write the model's styles" },
    );
  });
});
