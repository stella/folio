import { describe, expect, test } from "bun:test";

import type { Document, Paragraph, Run, Table } from "../../model/document";
import {
  applyDocumentOp,
  applyDocumentOps,
  getSourceReplayToken,
  inheritSourceReplayToken,
  normalizeForOps,
  registerSourceReplayDocument,
} from "../documentOps";
import { DOCUMENT_OP_TYPES, INHERIT_RUN_PROPS, OP_STORIES } from "../types";

const fixture = () => {
  const text = { type: "text", text: "ab" } as const;
  const formatting = { bold: true };
  const run = { type: "run", formatting, content: [text] } satisfies Run;
  const first = { type: "paragraph", paraId: "00000001", content: [run] } satisfies Paragraph;
  const second = {
    type: "paragraph",
    paraId: "00000002",
    content: [{ type: "run", content: [{ type: "text", text: "cd" }] }],
  } satisfies Paragraph;
  const document = { package: { document: { content: [first, second] } } } satisfies Document;
  return { document, first, second, run, formatting, text };
};

describe("source replay provenance", () => {
  test("identity belongs to the registered document, not its copied fields", () => {
    const { document } = fixture();
    const token = registerSourceReplayDocument(document);
    expect(registerSourceReplayDocument(document)).toBe(token);
    expect(getSourceReplayToken(document)).toBe(token);
    expect(getSourceReplayToken(structuredClone(document))).toBeUndefined();
    expect(getSourceReplayToken({ ...document })).toBeUndefined();
    const untracked = fixture().document;
    const target = { ...untracked };
    inheritSourceReplayToken(target, untracked);
    expect(getSourceReplayToken(target)).toBeUndefined();
    expect(Object.isFrozen(target.package.document.content)).toBe(false);
  });

  test("text edits inherit identity, freeze new graphs, and preserve unaffected blocks", () => {
    const { document, first, second } = fixture();
    const token = registerSourceReplayDocument(document);
    const applied = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      text: "X",
      runProps: INHERIT_RUN_PROPS,
    });
    if (applied.isErr()) throw applied.error;
    const next = applied.value.document;
    expect(getSourceReplayToken(next)).toBe(token);
    expect(next.package.document.content.at(0)).not.toBe(first);
    expect(next.package.document.content.at(1)).toBe(second);
    expect(Object.isFrozen(next.package.document.content)).toBe(true);
    const changed = next.package.document.content.at(0);
    if (changed?.type !== "paragraph") throw new TypeError("Expected paragraph");
    expect(Object.isFrozen(changed.content)).toBe(true);
    expect(() => changed.content.push({ type: "run", content: [] })).toThrow();
    const empty = applyDocumentOps(next, []);
    if (empty.isErr()) throw empty.error;
    expect(empty.value.document).toBe(next);
    expect(getSourceReplayToken(empty.value.document)).toBe(token);
  });

  test("tracked record descendants reject mutation while document roots stay extensible", () => {
    const { document, first, run, formatting, text } = fixture();
    registerSourceReplayDocument(document);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(run)).toBe(true);
    expect(Object.isFrozen(formatting)).toBe(true);
    expect(Object.isFrozen(text)).toBe(true);
    expect(() => Object.assign(text, { text: "changed" })).toThrow();
    expect(() => {
      formatting.bold = false;
    }).toThrow();
    expect(() => run.content.push(text)).toThrow();
    expect(Object.isFrozen(document)).toBe(false);
    expect(Object.isFrozen(document.package)).toBe(false);
    expect(Object.isFrozen(document.package.document)).toBe(false);
  });

  test("normalization carries identity into rebuilt blocks without mutating source", () => {
    const { document, first, second } = fixture();
    first.content.push({ type: "run", formatting: { bold: true }, content: [] });
    const token = registerSourceReplayDocument(document);
    const normalized = normalizeForOps(document);
    expect(getSourceReplayToken(normalized)).toBe(token);
    expect(normalized.package.document.content.at(0)).not.toBe(first);
    expect(normalized.package.document.content.at(1)).toBe(second);
    expect(first.content).toHaveLength(2);
    expect(Object.isFrozen(normalized.package.document.content.at(0))).toBe(true);
  });

  test("normalization preserves every already-normal nested block identity", () => {
    const { first, second } = fixture();
    const table = {
      type: "table",
      rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [first] }] }],
    } satisfies Table;
    const document = { package: { document: { content: [table, second] } } } satisfies Document;
    registerSourceReplayDocument(document);
    const normalized = normalizeForOps(document);
    expect(normalized.package.document.content).toBe(document.package.document.content);
    expect(normalized.package.document.content.at(0)).toBe(table);
    expect(normalizeForOps(normalized).package.document.content).toBe(
      document.package.document.content,
    );
  });

  test("background, section properties, styles and theme descendants are immutable", () => {
    const { document } = fixture();
    const background = { drawing: { rawXml: "<drawing/>" } };
    const finalSectionProperties = { pageWidth: 100, pageHeight: 200 };
    const styles = { styles: [], latentStyles: { count: 0 } };
    const theme = { fontScheme: { majorFont: { latin: "Arial" } } };
    const tracked = {
      package: {
        document: { ...document.package.document, background, finalSectionProperties },
        styles,
        theme,
      },
    } satisfies Document;
    registerSourceReplayDocument(tracked);
    expect(() => {
      background.drawing.rawXml = "changed";
    }).toThrow();
    expect(() => {
      finalSectionProperties.pageWidth = 300;
    }).toThrow();
    expect(() => {
      styles.latentStyles.count = 1;
    }).toThrow();
    expect(() => {
      theme.fontScheme.majorFont.latin = "Times New Roman";
    }).toThrow();
  });
});
