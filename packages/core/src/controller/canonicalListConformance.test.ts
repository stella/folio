import { expect, test } from "bun:test";
import { singletonManager } from "../prosemirror/schema";
import { shapeArrayBuffer, documentShape } from "../__tests__/documentShapes";
import JSZip from "jszip";

import {
  createCanonicalEditorHarness,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import {
  modelMarkdown,
  readBack,
  summarizeEffectiveParagraphs,
  summarizeState,
  parseShapeDocument,
  placeSelection,
} from "../__tests__/editorHarness";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "../docx/rezip";
import { findChildByNamespaceUri, parseXmlDocument } from "../docx/xmlParser";
import { parseDocx } from "../docx/parser";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { computeListRendering, getCachedNumberingMap } from "../docx/numberingParser";
import { paragraphNumberingReference } from "@stll/docx-core/model";
import { directParagraphIndentation } from "../prosemirror/paragraphIndentation";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { createHarnessState } from "../__tests__/editorHarness";

const CASES = (["editing", "suggesting"] as const).flatMap((mode) =>
  (["toggleBulletList", "toggleNumberedList"] as const).map((command) => ({ mode, command })),
);

test.each(CASES)(
  "canonical $command keeps list observations after save in $mode",
  async ({ mode, command }) => {
    const source = createEmptyDocument({ initialText: "List item" });
    const initial = source.package.document.content.at(0);
    if (initial?.type !== "paragraph") throw new TypeError("List fixture lost its paragraph.");
    initial.paraId = "12345678";
    const driver = createCanonicalEditorHarness(source, mode);
    try {
      driver.history.setSelection(3, 3);
      expect(driver.execute(driver.commandManager.requireCommand(command)())).toBe(true);
      expect(driver.refusals).toEqual([]);
      const model = driver.snapshot();
      const paragraph = model.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") throw new TypeError("List command lost its paragraph.");
      expect(paragraph.formatting?.numPr?.kind).toBe("reference");
      const bytes = await createDocx(model);
      const zip = await JSZip.loadAsync(bytes);
      const xml = await zip.file("word/document.xml")?.async("string");
      if (xml === undefined) throw new TypeError("Saved list fixture has no main story.");
      // The level owns indentation: canonical list selection must not author a paragraph w:ind.
      const root = parseXmlDocument(xml);
      if (root?.namespaceUri === undefined)
        throw new TypeError("Saved document lacks its Word namespace.");
      const namespaces = new Set([root.namespaceUri]);
      const body = findChildByNamespaceUri(root, namespaces, "body");
      const savedParagraph = findChildByNamespaceUri(body, namespaces, "p");
      const properties = findChildByNamespaceUri(savedParagraph, namespaces, "pPr");
      expect(findChildByNamespaceUri(properties, namespaces, "numPr")).not.toBeNull();
      expect(findChildByNamespaceUri(properties, namespaces, "ind")).toBeNull();
      const reopened = await parseDocx(bytes, { preloadFonts: false });
      expect(modelMarkdown(model)).toBe(modelMarkdown(reopened));
      const liveNode = driver.state.doc.firstChild;
      const reopenedNode = toProseDoc(reopened).firstChild;
      if (liveNode === null || reopenedNode === null)
        throw new TypeError("List projection lost its paragraph.");
      expect(directParagraphIndentation(expectParagraphAttrs(liveNode))).toBeUndefined();
      expect(directParagraphIndentation(expectParagraphAttrs(reopenedNode))).toBeUndefined();
      const back = await readBack(new Uint8Array(bytes));
      expect(summarizeEffectiveParagraphs(driver.state)).toEqual(back.effective);
    } finally {
      driver.dispose();
    }
  },
);

const CACHE_CASES = (["bullet", "numbered"] as const).flatMap((kind) =>
  (["missing", "stale", "current"] as const).map((cache) => ({ kind, cache })),
);

test.each(CACHE_CASES)(
  "canonical $kind list reads authoritative definitions with $cache cache",
  async ({ kind, cache }) => {
    const source = fromMarkdown(kind === "bullet" ? "- List item" : "1. List item");
    const paragraph = source.package.document.content.at(0);
    const numbering = source.package.numbering;
    const level = numbering?.abstractNums.at(0)?.levels.at(0);
    if (paragraph?.type !== "paragraph" || numbering === undefined || level === undefined) {
      throw new TypeError("Existing-list fixture lacks numbering.");
    }
    paragraph.paraId = "12345678";
    // The authored paragraph has no indentation; the numbering level owns it.
    const numPr = paragraph.formatting?.numPr;
    if (numPr?.kind !== "reference")
      throw new TypeError("Existing-list fixture lacks a reference.");
    paragraph.formatting = {
      numPr: paragraphNumberingReference({ numId: numPr.numId, ilvl: numPr.ilvl }),
    };
    level.pPr = { indentLeft: 1440, indentFirstLine: -480, hangingIndent: true };
    if (cache === "missing") Reflect.deleteProperty(paragraph, "listRendering");
    if (cache === "stale") {
      level.lvlText = kind === "bullet" ? "◆" : "%1)";
      level.start = 7;
    }
    if (cache === "current") {
      const rendering = computeListRendering(numPr, getCachedNumberingMap(numbering));
      if (rendering === null)
        throw new TypeError("Existing-list fixture cannot resolve its rendering.");
      paragraph.listRendering = rendering;
    }
    const authored = structuredClone(source);
    const projected = toProseDoc(source);
    const attrs = projected.firstChild?.attrs;
    expect(attrs?.indentLeft).toBe(1440);
    expect(attrs?.indentFirstLine).toBe(-480);
    expect(attrs?.hangingIndent).toBe(true);
    const bytes = await createDocx(source);
    const reopened = await parseDocx(bytes, { preloadFonts: false });
    expect(modelMarkdown(source)).toBe(modelMarkdown(reopened));
    expect(summarizeEffectiveParagraphs(createHarnessState(source, "editing"))).toEqual(
      summarizeEffectiveParagraphs(createHarnessState(reopened, "editing")),
    );
    expect(source).toEqual(authored);
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")?.async("string");
    if (xml === undefined) throw new TypeError("Saved existing-list fixture has no main story.");
    const root = parseXmlDocument(xml);
    if (root?.namespaceUri === undefined)
      throw new TypeError("Saved document lacks its Word namespace.");
    const namespaces = new Set([root.namespaceUri]);
    const properties = findChildByNamespaceUri(
      findChildByNamespaceUri(findChildByNamespaceUri(root, namespaces, "body"), namespaces, "p"),
      namespaces,
      "pPr",
    );
    expect(findChildByNamespaceUri(properties, namespaces, "ind")).toBeNull();
  },
);

test.each(["inherited", "explicit", "edited"] as const)(
  "parsed list keeps $0 indentation provenance",
  async (kind) => {
    const source = fromMarkdown("1. List item");
    const paragraph = source.package.document.content.at(0);
    const numbering = source.package.numbering;
    const level = numbering?.abstractNums.at(0)?.levels.at(0);
    if (paragraph?.type !== "paragraph" || level === undefined)
      throw new TypeError("Indent fixture lacks numbering.");
    paragraph.paraId = "12345678";
    paragraph.formatting = {
      numPr: paragraph.formatting?.numPr,
      ...(kind === "explicit" ? { indentLeft: 900, indentFirstLine: 0 } : {}),
    };
    level.pPr = { indentLeft: 1440, indentFirstLine: -480, hangingIndent: true };
    const parsed = await parseDocx(await createDocx(source), { preloadFonts: false });
    const parsedParagraph = parsed.package.document.content.at(0);
    if (parsedParagraph?.type !== "paragraph")
      throw new TypeError("Parsed indent fixture lost paragraph.");
    if (kind === "edited") {
      parsedParagraph.formatting = {
        ...parsedParagraph.formatting,
        indentLeft: 901,
        indentFirstLine: 0,
        hangingIndent: false,
      };
    }
    const node = toProseDoc(parsed).firstChild;
    if (node === null) throw new TypeError("Indent projection lost its paragraph.");
    const direct = directParagraphIndentation(expectParagraphAttrs(node));
    expect(direct).toEqual(
      kind === "inherited"
        ? undefined
        : {
            indentLeft: kind === "explicit" ? 900 : 901,
            indentFirstLine: 0,
            hangingIndent: false,
          },
    );
  },
);

const INAPPLICABLE_CASES = (["editing", "suggesting"] as const).flatMap((mode) =>
  (
    ["restartNumbering", "continueNumbering", "increaseListLevel", "decreaseListLevel"] as const
  ).map((command) => ({ mode, command })),
);

test.each(INAPPLICABLE_CASES)(
  "canonical $command on prose is an atomic no-change in $mode",
  ({ mode, command }) => {
    const source = createEmptyDocument({ initialText: "Plain paragraph" });
    const paragraph = source.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new TypeError("No-change fixture lost paragraph.");
    paragraph.paraId = "12345678";
    const driver = createCanonicalEditorHarness(source, mode);
    try {
      driver.history.setSelection(3, 3);
      const before = driver.snapshot();
      const projection = driver.state.doc.toJSON();
      const selection = driver.state.selection.toJSON();
      expect(driver.execute(driver.commandManager.requireCommand(command)())).toBe(false);
      expect(driver.refusals).toEqual([]);
      expect(driver.snapshot()).toEqual(before);
      expect(driver.state.doc.toJSON()).toEqual(projection);
      expect(driver.state.selection.toJSON()).toEqual(selection);
      expect(driver.history.canUndo()).toBe(false);
      expect(driver.history.canRedo()).toBe(false);
    } finally {
      driver.dispose();
    }
  },
);

// The original list oracle compared effective layout but omitted reviewer direct formatting.
test.each(["editing", "suggesting"] as const)(
  "unrelated list formatting preserves authored indentation after save in %s",
  async (mode) => {
    const source = await parseShapeDocument(
      new Uint8Array(await shapeArrayBuffer("single-decimal-list")),
    );
    const driver = createCanonicalEditorHarness(source, mode);
    try {
      const positions: number[] = [];
      driver.state.doc.forEach((node, offset) => {
        if (node.attrs.numPr?.kind === "reference") positions.push(offset + 1);
      });
      for (const position of positions) {
        driver.history.setSelection(position, position);
        expect(driver.execute(driver.commandManager.requireCommand("setAlignment")("center"))).toBe(
          true,
        );
        const live = summarizeState(driver.state);
        const back = await readBack((await saveCanonicalHarnessDocument(driver.snapshot())).bytes);
        expect(back.summary).toEqual(live);
        expect(back.effective).toEqual(summarizeEffectiveParagraphs(driver.state));
      }
    } finally {
      driver.dispose();
    }
  },
);

const PROVENANCE_COMMANDS = [
  { id: "hanging", create: () => singletonManager.requireCommand("setIndentFirstLine")(360, true) },
  {
    id: "first-line",
    create: () => singletonManager.requireCommand("setIndentFirstLine")(360, false),
  },
  { id: "level-up", create: () => singletonManager.requireCommand("increaseListLevel")() },
  { id: "level-down", create: () => singletonManager.requireCommand("decreaseListLevel")() },
  { id: "remove-list", create: () => singletonManager.requireCommand("removeList")() },
  { id: "style", create: () => singletonManager.requireCommand("applyStyle")("Heading2") },
  { id: "clear-style", create: () => singletonManager.requireCommand("clearStyle")() },
];
const PROVENANCE_CASES = (["editing", "suggesting"] as const).flatMap((mode) =>
  ["single-decimal-list", "single-bullet-list", "outline-level-numbered"].flatMap((shape) =>
    PROVENANCE_COMMANDS.map((command) => ({ mode, shape, id: command.id, create: command.create })),
  ),
);

test.each(PROVENANCE_CASES)(
  "$shape $id preserves direct and effective indentation after save in $mode",
  async ({ mode, shape, create, id }) => {
    const source = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer(shape)));
    const driver = createCanonicalEditorHarness(source, mode);
    try {
      let position: number | undefined;
      driver.state.doc.forEach((node, offset) => {
        if (node.attrs.numPr?.kind === "reference") position = offset + 1;
      });
      if (position === undefined) throw new TypeError("Provenance fixture lacks a numbered item.");
      driver.history.setSelection(position, position);
      if (id === "clear-style") {
        expect(driver.execute(singletonManager.requireCommand("applyStyle")("Heading2"))).toBe(
          true,
        );
      }
      expect(driver.execute(create())).toBe(true);
      const back = await readBack((await saveCanonicalHarnessDocument(driver.snapshot())).bytes);
      expect(back.summary).toEqual(summarizeState(driver.state));
      expect(back.effective).toEqual(summarizeEffectiveParagraphs(driver.state));
    } finally {
      driver.dispose();
    }
  },
);

const JOIN_CASES = (["editing", "suggesting"] as const).flatMap((mode) =>
  ["mixed-lists", "style-numbered-headings"].map((shape) => ({ mode, shape })),
);

test.each(JOIN_CASES)(
  "$shape joined paragraphs derive their list from current numbering in $mode",
  async ({ mode, shape }) => {
    const source = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer(shape)));
    const driver = createCanonicalEditorHarness(source, mode);
    try {
      const selected = placeSelection(driver.state, documentShape(shape).focus, "cross-paragraph");
      if (!selected) throw new TypeError("List join fixture has no cross-paragraph selection.");
      driver.dispatch(driver.state.tr.setSelection(selected.selection));
      driver.cut();
      expect(driver.refusals).toEqual([]);
      const model = driver.snapshot();
      const back = await readBack((await saveCanonicalHarnessDocument(model)).bytes);
      expect(back.summary).toEqual(summarizeState(driver.state));
      expect(back.effective).toEqual(summarizeEffectiveParagraphs(driver.state));
      expect(back.markdown).toBe(modelMarkdown(model));
    } finally {
      driver.dispose();
    }
  },
);
