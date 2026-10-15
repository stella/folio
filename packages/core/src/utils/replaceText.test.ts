import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createDefaultFindOptions, findInDocument } from "./findReplace";
import { createEmptyDocument } from "./createDocument";
import { replaceTextInDocument } from "./replaceText";
import type { Paragraph, ParagraphContent } from "../types/document";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const documentWithRevision = () => {
  const template = createEmptyDocument();
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "C0000001",
    formatting: { alignment: "center" },
    content: [
      run("left"),
      {
        type: "deletion",
        info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
        content: [run(" hidden ")],
      },
      run("right"),
    ],
  };
  return {
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content: [paragraph] },
    },
  };
};

type RevisionKind = Extract<ParagraphContent, { info: unknown }>["type"];

const revisionWrapperFactories = {
  insertion: (text: string): ParagraphContent => ({
    type: "insertion",
    info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
    content: [run(text)],
  }),
  deletion: (text: string): ParagraphContent => ({
    type: "deletion",
    info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
    content: [run(text)],
  }),
  moveFrom: (text: string): ParagraphContent => ({
    type: "moveFrom",
    info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
    content: [run(text)],
  }),
  moveTo: (text: string): ParagraphContent => ({
    type: "moveTo",
    info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
    content: [run(text)],
  }),
} as const satisfies Record<RevisionKind, (text: string) => ParagraphContent>;

const revisionCases = Object.entries(revisionWrapperFactories).flatMap(([kind, wrap]) =>
  ["stock", "clause", "review"].map((needle) => ({ kind, wrap, needle })),
);

const documentWithContent = (content: ParagraphContent[]) => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph" as const, paraId: "C0000001", content }],
      },
    },
  };
};

describe("replaceTextInDocument", () => {
  test("resolves replacement offsets through tracked revision wrappers", () => {
    const source = documentWithRevision();
    const originalContent = structuredClone(source.package.document.content);
    const replaced = replaceTextInDocument(
      source,
      { start: { paragraphIndex: 0, offset: 0 }, end: { paragraphIndex: 0, offset: 4 } },
      "replacement",
    );
    const paragraph = replaced.package.document.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") {
      panic("Expected the replaced paragraph");
    }
    expect(toProseDoc(replaced).textContent).toBe("replacement hidden right");
    expect(paragraph.content.some((item) => item.type === "deletion")).toBe(true);
    expect(source.package.document.content).toEqual(originalContent);
    expect(paragraph.formatting).toEqual({ alignment: "center" });
  });

  test.each(revisionCases)("replaces text inside $kind revision", ({ kind, wrap, needle }) => {
    const original = documentWithContent([
      run("prefix "),
      wrap(`before ${needle} after`),
      run(" suffix"),
    ]);
    const originalContent = structuredClone(original.package.document.content);
    const sourceText = toProseDoc(original).textContent;
    const match = findInDocument(original, needle, createDefaultFindOptions()).at(0);
    if (match === undefined) {
      throw new Error(`Expected to find ${needle} in ${kind}`);
    }
    const replaced = replaceTextInDocument(
      original,
      {
        start: { paragraphIndex: match.paragraphIndex, offset: match.startOffset },
        end: { paragraphIndex: match.paragraphIndex, offset: match.endOffset },
      },
      "shares",
    );

    expect(toProseDoc(original).textContent).toBe(sourceText);
    expect(original.package.document.content).toEqual(originalContent);
    expect(toProseDoc(replaced).textContent).toBe("prefix before shares after suffix");
    const paragraph = replaced.package.document.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") {
      panic("Expected the replaced paragraph");
    }
    expect(paragraph.content.at(1)?.type).toBe(kind);
  });

  test("applies requested formatting to replacement text", () => {
    const replaced = replaceTextInDocument(
      documentWithRevision(),
      { start: { paragraphIndex: 0, offset: 0 }, end: { paragraphIndex: 0, offset: 4 } },
      "replacement",
      { bold: true },
    );
    const paragraph = replaced.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") {
      panic("Expected the replaced paragraph");
    }
    const insertedRun = paragraph.content.at(0);
    expect(insertedRun?.type).toBe("run");
    if (insertedRun?.type !== "run") {
      panic("Expected replacement text in a run");
    }
    expect(insertedRun.formatting?.bold).toBe(true);
  });
});
