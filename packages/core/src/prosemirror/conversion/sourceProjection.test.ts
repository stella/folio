import { describe, expect, test } from "bun:test";
import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import { registerSourceReplayDocument } from "@stll/docx-core/ops";

import {
  assignDocumentParagraphPropertySourceContract,
  copyDocumentParagraphPropertySourceContract,
  copyParagraphPropertySource,
} from "../../docx/paragraphPropertySource";
import type { Document, Paragraph } from "../../types/document";
import { currentSourceProjection, documentProjectionInput } from "./sourceProjection";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc, type ToProseDocOptions } from "./toProseDoc";

type ProjectionVersionKey = keyof ReturnType<typeof documentProjectionInput> | "tokens";
type VersionMutation = (document: Document) => void;
type VersionMutationCase = {
  prepare: (document: Document) => void;
  mutate: VersionMutation;
};

const paragraph = (): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text: "source" }] }],
});

type SourceDocumentOptions = {
  sourceReplay?: "tracked" | "untracked";
  prepare?: (document: Document) => void;
};

const sourceDocument = ({
  sourceReplay = "tracked",
  prepare,
}: SourceDocumentOptions = {}): Document => {
  const document: Document = { package: { document: { content: [paragraph()] } } };
  assignDocumentParagraphPropertySourceContract(document, "a".repeat(64));
  // Ordinary immutable derivations carry a replaceable provenance binding.
  // The factory's original binding deliberately cannot be overwritten.
  const result = { ...document };
  prepare?.(result);
  if (sourceReplay === "tracked") registerSourceReplayDocument(result);
  return result;
};

const versionMutations = {
  content: {
    prepare: () => undefined,
    mutate: (document) => {
      const sourceParagraph = document.package.document.content.at(0);
      if (sourceParagraph?.type !== "paragraph") {
        throw new TypeError("Expected source paragraph");
      }
      sourceParagraph.content.push({ type: "run", content: [{ type: "text", text: " nested" }] });
    },
  },
  styles: {
    prepare: (document) => {
      document.package.styles = { docDefaults: { rPr: { bold: true } }, styles: [] };
    },
    mutate: (document) => {
      const styles = document.package.styles;
      if (!styles?.docDefaults?.rPr) throw new TypeError("Expected document defaults");
      styles.docDefaults.rPr.italic = true;
    },
  },
  theme: {
    prepare: (document) => {
      document.package.theme = { colorScheme: { accent1: "112233" } };
    },
    mutate: (document) => {
      const colorScheme = document.package.theme?.colorScheme;
      if (!colorScheme) throw new TypeError("Expected theme color scheme");
      colorScheme.accent1 = "445566";
    },
  },
  finalSectionStart: {
    prepare: (document) => {
      document.package.document.sections = [
        { properties: { sectionStart: "continuous" }, content: [] },
      ];
    },
    mutate: (document) => {
      const finalSection = document.package.document.sections?.at(-1);
      if (!finalSection) throw new TypeError("Expected final section");
      finalSection.properties.sectionStart = "nextPage";
    },
  },
  adjustLineHeightInTable: {
    prepare: (document) => {
      document.package.settings = { adjustLineHeightInTable: true };
    },
    mutate: (document) => {
      document.package.settings = {};
    },
  },
  doNotUseIndentAsNumberingTabStop: {
    prepare: (document) => {
      document.package.settings = { doNotUseIndentAsNumberingTabStop: true };
    },
    mutate: (document) => {
      document.package.settings = {};
    },
  },
  contract: {
    prepare: () => undefined,
    mutate: (document) => {
      const other = sourceDocument({ sourceReplay: "untracked" });
      assignDocumentParagraphPropertySourceContract(other, "b".repeat(64));
      copyDocumentParagraphPropertySourceContract(document, other);
    },
  },
  tokens: {
    prepare: () => undefined,
    mutate: (document) => {
      const other = sourceDocument({ sourceReplay: "untracked" });
      assignDocumentParagraphPropertySourceContract(other, "c".repeat(64));
      const target = document.package.document.content.at(0);
      const source = other.package.document.content.at(0);
      if (target?.type !== "paragraph" || source?.type !== "paragraph") {
        throw new TypeError("Expected bound paragraphs");
      }
      copyParagraphPropertySource(target, source);
    },
  },
} satisfies Record<ProjectionVersionKey, VersionMutationCase>;

describe("source projection cache versions", () => {
  test("unchanged projections reuse the same ProseMirror document", () => {
    const document = sourceDocument();
    const firstProjection = toProseDoc(document);

    expect(toProseDoc(document)).toBe(firstProjection);
    expect(currentSourceProjection(document)).toBe(firstProjection);
  });

  test.each(Object.entries(versionMutations))("guards tracked %s changes", (key, change) => {
    const document = sourceDocument({ prepare: change.prepare });
    const originalProjection = toProseDoc(document);

    expect(currentSourceProjection(document)).toBe(originalProjection);
    if (key === "content" || key === "styles" || key === "theme" || key === "tokens") {
      expect(() => change.mutate(document)).toThrow();
      expect(currentSourceProjection(document)).toBe(originalProjection);
      return;
    }
    change.mutate(document);
    expect(currentSourceProjection(document)).toBeUndefined();

    const updatedProjection = toProseDoc(document);
    expect(updatedProjection).not.toBe(originalProjection);
    expect(currentSourceProjection(document)).toBe(updatedProjection);
  });

  test.each(Object.entries(versionMutations))(
    "untracked %s changes never authorize projection reuse",
    (key, change) => {
      const document = sourceDocument({ sourceReplay: "untracked", prepare: change.prepare });
      const projected = toProseDoc(document);
      fromProseDoc(projected, document);
      fromProseDoc(projected, document);
      expect(currentSourceProjection(document)).toBeUndefined();

      change.mutate(document);
      expect(currentSourceProjection(document)).toBeUndefined();
      if (key === "contract" || key === "tokens") {
        expect(() => fromProseDoc(projected, document)).toThrow();
        return;
      }
      expect(fromProseDoc(projected, document).package.document.content).toEqual(
        fromProseDoc(projected, document, { reuse: "none" }).package.document.content,
      );
    },
  );
});

test("warned projections still deliver warnings on each conversion", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "run",
                content: [{ type: "softHyphen" }, { type: "break", breakType: "page" }],
              },
            ],
          },
        ],
      },
    },
  };
  const details: (string | undefined)[] = [];
  const warn: NonNullable<ToProseDocOptions["warn"]> = ({ code, detail }) => {
    if (code === PARSE_WARNING_CODES.pageBreakProjectionApproximated) details.push(detail);
  };

  toProseDoc(document, { warn });
  toProseDoc(document, { warn });

  expect(details).toHaveLength(2);
  expect(details.at(0)).toContain("A page-break-bearing run also holds softHyphen");
});

test("invalidates when content is replaced", () => {
  const document = sourceDocument();
  const originalProjection = toProseDoc(document);
  expect(currentSourceProjection(document)).toBe(originalProjection);

  document.package.document.content = [paragraph()];
  expect(currentSourceProjection(document)).toBeUndefined();

  const updatedProjection = toProseDoc(document);
  expect(updatedProjection).not.toBe(originalProjection);
  expect(currentSourceProjection(document)).toBe(updatedProjection);
});
