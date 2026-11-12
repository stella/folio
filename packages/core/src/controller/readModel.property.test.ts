/**
 * `readModel` hands a host's document model to readers that take a plain
 * `Document`. The model is shared with the session, its caches and the host's
 * undo history, so this file runs every reader a host calls through
 * `readModel` over deep-frozen documents: none may throw on a frozen document,
 * and each must return what it returns for the same document unfrozen.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

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
import { resolveSectionHeaderFooterRefs } from "../layout-engine/headerFooterRefs";
import { TableSelectionManager } from "../managers/TableSelectionManager";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { findInDisplayedDocument } from "../prosemirror/findReplaceSelection";
import { singletonManager } from "../prosemirror/schema";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import type { Document } from "../types/document";
import { readModel } from "../types/documentModel";
import { createFindSearchKey } from "../utils/findSearchKey";
import {
  resolveEffectiveSectionProperties,
  resolveHeaderFooterContent,
} from "../utils/headerFooter";
import { getDocumentWatermark } from "../watermark";
import { createCanonicalSession } from "./canonicalSession";
import { createHiddenEditorState } from "./hiddenEditorManager";

setDefaultTimeout(propertyTestTimeout(120_000));

const MAX_LEVEL = 8;

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

type Query = {
  searchText: string;
  matchCase: boolean;
  tableIndex: number;
  rowIndex: number;
  columnIndex: number;
};

/** Every reader a host calls with its model, and what each returns. */
const readEverything = (document: Document, query: Query) => {
  const styles = document.package.styles;
  const resolver = createStyleResolver(styles);
  const styleIds = styles?.styles.map(({ styleId }) => styleId) ?? [];
  const numbering = document.package.numbering;
  const numberingMap = numbering ? getCachedNumberingMap(numbering) : null;
  const numIds = numbering?.nums.map(({ numId }) => numId) ?? [];
  const state = createHiddenEditorState({
    document,
    styles: styles ?? null,
    manager: singletonManager,
  });
  const options = { matchCase: query.matchCase, matchWholeWord: false };
  const headerFooter = resolveHeaderFooterContent(document.package);
  return {
    styles: styleIds.map((styleId) => ({
      style: resolver.getStyle(styleId),
      paragraph: resolver.resolveParagraphStyle(styleId),
      inTable: resolver.resolveParagraphStyleInTable(styleId, undefined),
      run: resolver.resolveRunStyle(styleId),
      own: resolver.getRunStyleOwnProperties(styleId),
      next: resolver.getNextStyleId(styleId),
      has: resolver.hasStyle(styleId),
    })),
    paragraphStyles: resolver.getParagraphStyles(),
    tableStyles: resolver.getTableStyles(),
    docDefaults: resolver.getDocDefaults(),
    defaultParagraphStyle: resolver.getDefaultParagraphStyle(),
    defaultCharacterStyle: resolver.getDefaultCharacterStyle(),
    defaultTableStyle: resolver.getDefaultTableStyle(),
    numbering: numIds.map((numId) => ({
      has: numberingMap?.hasNumbering(numId),
      instance: numberingMap?.getInstance(numId),
      abstractNumId: numberingMap?.getAbstractNumId(numId),
      abstract: numberingMap?.getAbstract(
        numberingMap.getAbstractNumId(numId) ?? -1,
      ),
      levels: Array.from({ length: MAX_LEVEL + 1 }, (_, ilvl) =>
        numberingMap?.getLevel(numId, ilvl),
      ),
    })),
    hiddenEditor: state.doc.toJSON(),
    built: fromProseDoc(state.doc, document, {
      stylesheetSource: { type: "package" },
    }),
    find: findInDisplayedDocument({
      document,
      searchText: query.searchText,
      options,
    }),
    findKey: createFindSearchKey({
      query: query.searchText,
      options,
      document,
    }),
    table: new TableSelectionManager().selectCell(document, query),
    headerFooter,
    sectionProperties: resolveEffectiveSectionProperties(
      document.package.document,
      headerFooter.hasTitlePg,
    ),
    sectionRefs: resolveSectionHeaderFooterRefs(document),
    watermark: getDocumentWatermark(document),
  };
};

const queryArbitrary = fc.record({
  searchText: fc.constantFrom("", "a", "the", "Item", "e", " ", "Note"),
  matchCase: fc.boolean(),
  tableIndex: fc.nat({ max: 2 }),
  rowIndex: fc.nat({ max: 2 }),
  columnIndex: fc.nat({ max: 2 }),
});

describe("readModel", () => {
  test("a frozen document refuses a write", async () => {
    const [{ document }] = await loadShapes();
    const frozen = deepFreeze(
      cloneDocumentWithParagraphPropertySources(document),
    );
    expect(() => {
      (readModel(frozen).package.document.content as unknown[]).push(null);
    }).toThrow(TypeError);
  });

  test("every reader a host calls reads a frozen model as it reads a plain one", async () => {
    const shapes = await loadShapes();
    await assertProperty(
      fc.property(
        fc.constantFrom(...shapes),
        queryArbitrary,
        ({ document }, query) => {
          const frozen = deepFreeze(
            cloneDocumentWithParagraphPropertySources(document),
          );
          const plain = cloneDocumentWithParagraphPropertySources(document);
          expect(readEverything(readModel(frozen), query)).toEqual(
            readEverything(plain, query),
          );
        },
      ),
      { id: "host readers never write the model" },
    );
  });

  test("every reader a host calls reads a canonical session's own document frozen as it reads it plain", async () => {
    const shapes = await loadShapes();
    const sessionDocument = (document: Document) => {
      const session = createCanonicalSession(
        cloneDocumentWithParagraphPropertySources(document),
      );
      return session.isOk() ? session.value.document : null;
    };
    const seeded = shapes.filter(({ document }) => sessionDocument(document));
    expect(seeded.length).toBeGreaterThan(0);
    await assertProperty(
      fc.property(
        fc.constantFrom(...seeded),
        queryArbitrary,
        ({ document }, query) => {
          const frozen = sessionDocument(document);
          const plain = sessionDocument(document);
          if (!frozen || !plain) throw new Error("A seeded shape refused.");
          deepFreeze(frozen);
          expect(readEverything(readModel(frozen), query)).toEqual(
            readEverything(readModel(plain), query),
          );
        },
      ),
      { id: "host readers never write the canonical session's document" },
    );
  });
});
