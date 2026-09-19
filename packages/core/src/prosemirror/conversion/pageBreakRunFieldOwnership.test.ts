import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import type { Hyperlink, Run, RunContent } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { schema } from "../schema";
import { validateProseMirrorDocument } from "../validation";
import { fromProseDoc } from "./fromProseDoc";
import { type ToProseDocOptions, toProseDoc } from "./toProseDoc";

const REVISION = {
  id: 91,
  author: "Reviewer",
  date: "2026-09-09T00:00:00.000Z",
} as const;

const APPROXIMATED_FIELD_RESULT_CONTENT = {
  fieldChar: { type: "fieldChar", charType: "begin" },
  instrText: { type: "instrText", text: " PAGE " },
  noBreakHyphen: { type: "noBreakHyphen" },
  softHyphen: { type: "softHyphen" },
  "text-box shape": {
    type: "shape",
    shape: {
      type: "shape",
      shapeType: "rect",
      size: { width: 914_400, height: 457_200 },
      textBody: { content: [] },
    },
  },
} as const satisfies Record<string, RunContent>;

const approximatedFieldResultDetail = (content: RunContent): string =>
  `A field result with an explicit page break also holds ${
    content.type === "shape" ? "a text-box shape" : content.type
  }`;

type ProjectionWarning = { code: string; detail?: string };

/**
 * A page break folio lays out approximately opens, saves and reports.
 *
 * Folio used to refuse these documents. Opening a file Word opens is not
 * optional, so the check is now that the conversion completes and says on the
 * parse-warning channel what it approximated.
 */
const expectProjectionWarning = (
  convert: (options: ToProseDocOptions) => unknown,
  detail: string,
): void => {
  const warnings: ProjectionWarning[] = [];
  convert({
    warn: (report) => {
      warnings.push({
        code: report.code,
        ...(report.detail === undefined ? {} : { detail: report.detail }),
      });
    },
  });
  expect(warnings).toContainEqual({
    code: PARSE_WARNING_CODES.pageBreakProjectionApproximated,
    detail,
  });
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/document.xml");
  if (!file) {
    throw new Error("Missing word/document.xml");
  }
  return file.async("text");
};

describe("page-break run field ownership", () => {
  test.each(["simpleField", "complexField"] as const)(
    "preserves a page break inside a tracked %s result",
    (fieldType) => {
      const source = createEmptyDocument();
      const run = {
        type: "run" as const,
        formatting: { underline: { style: "single" as const } },
        content: [
          { type: "text" as const, text: "A" },
          { type: "break" as const, breakType: "page" as const, clear: "all" as const },
          { type: "text" as const, text: "B" },
        ],
      };
      const field =
        fieldType === "simpleField"
          ? {
              type: "simpleField" as const,
              instruction: "REF target",
              fieldType: "REF" as const,
              content: [run],
            }
          : {
              type: "complexField" as const,
              instruction: "REF target",
              fieldType: "REF" as const,
              fieldCode: [],
              fieldResult: [run],
            };
      source.package.document.content = [
        {
          type: "paragraph",
          content: [{ type: "insertion", info: REVISION, content: [field] }],
        },
      ];

      const prose = toProseDoc(source);
      const structuredField = prose.firstChild?.firstChild;
      expect(structuredField?.type.name).toBe("structuredField");
      expect(structuredField?.marks.some((mark) => mark.type.name === "insertion")).toBe(true);
      expect(structuredField?.child(1).type.name).toBe("pageBreakRun");

      const roundTripped = fromProseDoc(prose, source);
      const paragraph = roundTripped.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") {
        throw new Error("Expected paragraph");
      }
      const insertion = paragraph.content.at(0);
      if (insertion?.type !== "insertion") {
        throw new Error("Expected insertion");
      }
      const restoredField = insertion.content.at(0);
      if (restoredField?.type !== fieldType) {
        throw new Error(`Expected ${fieldType}`);
      }
      const resultRun =
        restoredField.type === "simpleField"
          ? restoredField.content.at(0)
          : restoredField.fieldResult.at(0);
      expect(resultRun?.type).toBe("run");
      if (resultRun?.type === "run") {
        expect(resultRun.content).toEqual(run.content);
      }
    },
  );

  test.each(["simpleField", "complexField"] as const)(
    "keeps a page-break-only %s result at a save/reopen fixed point",
    async (fieldType) => {
      const source = createEmptyDocument();
      const resultRun = {
        type: "run" as const,
        formatting: { italic: true },
        content: [{ type: "break" as const, breakType: "page" as const, clear: "left" as const }],
      };
      const field =
        fieldType === "simpleField"
          ? {
              type: "simpleField" as const,
              instruction: "REF target",
              fieldType: "REF" as const,
              content: [resultRun],
            }
          : {
              type: "complexField" as const,
              instruction: "REF target",
              fieldType: "REF" as const,
              fieldCode: [],
              fieldResult: [resultRun],
            };
      source.package.document.content = [{ type: "paragraph", content: [field] }];

      const firstModel = fromProseDoc(toProseDoc(source), source);
      const firstBuffer = await createDocx(firstModel);
      const reopened = await parseDocx(firstBuffer);
      const secondModel = fromProseDoc(toProseDoc(reopened), reopened);
      const secondBuffer = await createDocx(secondModel);

      expect(await documentXml(secondBuffer)).toBe(await documentXml(firstBuffer));
      const paragraph = secondModel.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") {
        throw new Error("Expected paragraph");
      }
      const restoredField = paragraph.content.at(0);
      if (restoredField?.type !== fieldType) {
        throw new Error(`Expected ${fieldType}`);
      }
      const restoredRuns =
        restoredField.type === "simpleField"
          ? restoredField.content.filter((content) => content.type === "run")
          : restoredField.fieldResult;
      expect(restoredRuns).toHaveLength(1);
      expect(restoredRuns.at(0)?.content).toEqual([
        { type: "break", breakType: "page", clear: "left" },
      ]);
    },
  );

  for (const fieldKind of ["simpleField", "complexField"] as const) {
    for (const ownership of ["direct", "tracked"] as const) {
      test.each(Object.entries(APPROXIMATED_FIELD_RESULT_CONTENT))(
        `reports an approximation for %s in a sibling ${ownership} ${fieldKind} result run`,
        (_, unsupportedContent) => {
          const source = createEmptyDocument();
          const pageBreakRun: Run = {
            type: "run",
            content: [{ type: "break", breakType: "page" }],
          };
          const unsupportedRun: Run = { type: "run", content: [unsupportedContent] };
          const field =
            fieldKind === "simpleField"
              ? {
                  type: "simpleField" as const,
                  instruction: "REF target",
                  fieldType: "REF" as const,
                  content: [pageBreakRun, unsupportedRun],
                }
              : {
                  type: "complexField" as const,
                  instruction: "REF target",
                  fieldType: "REF" as const,
                  fieldCode: [],
                  fieldResult: [pageBreakRun, unsupportedRun],
                };
          source.package.document.content = [
            {
              type: "paragraph",
              content:
                ownership === "direct"
                  ? [field]
                  : [{ type: "insertion", info: REVISION, content: [field] }],
            },
          ];

          expectProjectionWarning(
            (options) => toProseDoc(source, options),
            approximatedFieldResultDetail(unsupportedContent),
          );
        },
      );
    }
  }

  for (const ownership of ["direct", "tracked"] as const) {
    test.each(Object.entries(APPROXIMATED_FIELD_RESULT_CONTENT))(
      `reports an approximation for %s in a sibling ${ownership} simple-field hyperlink run`,
      (_, unsupportedContent) => {
        const source = createEmptyDocument();
        const hyperlink: Hyperlink = {
          type: "hyperlink",
          href: "https://example.test/field-result",
          children: [{ type: "run", content: [unsupportedContent] }],
        };
        const field = {
          type: "simpleField" as const,
          instruction: "REF target",
          fieldType: "REF" as const,
          content: [
            {
              type: "run" as const,
              content: [{ type: "break" as const, breakType: "page" as const }],
            },
            hyperlink,
          ],
        };
        source.package.document.content = [
          {
            type: "paragraph",
            content:
              ownership === "direct"
                ? [field]
                : [{ type: "insertion", info: REVISION, content: [field] }],
          },
        ];

        expectProjectionWarning(
          (options) => toProseDoc(source, options),
          approximatedFieldResultDetail(unsupportedContent),
        );
      },
    );
  }

  test.each(Object.entries(APPROXIMATED_FIELD_RESULT_CONTENT))(
    "reports an approximation for %s beside a top-level page break",
    (_, unsupportedContent) => {
      const source = createEmptyDocument();
      source.package.document.content = [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [{ type: "break", breakType: "page" }, unsupportedContent],
            },
          ],
        },
      ];

      expectProjectionWarning(
        (options) => toProseDoc(source, options),
        `A page-break-bearing run also holds ${
          unsupportedContent.type === "shape" ? "a text-box shape" : unsupportedContent.type
        }`,
      );
    },
  );

  test.each(["direct", "tracked"] as const)(
    "fails closed for a page break in a %s complex-field instruction",
    (ownership) => {
      const source = createEmptyDocument();
      const field = {
        type: "complexField" as const,
        instruction: "REF target",
        fieldType: "REF" as const,
        fieldCode: [
          {
            type: "run" as const,
            content: [{ type: "break" as const, breakType: "page" as const }],
          },
        ],
        fieldResult: [],
      };
      source.package.document.content = [
        {
          type: "paragraph",
          content:
            ownership === "direct"
              ? [field]
              : [{ type: "insertion", info: REVISION, content: [field] }],
        },
      ];

      expectProjectionWarning(
        (options) => toProseDoc(source, options),
        "A complex-field instruction holds an explicit page break",
      );
    },
  );

  test.each(["hyperlink", "bookmark", "text-box anchor"] as const)(
    "fails closed before export when a complex page-break result also contains a %s",
    (unsupportedChild) => {
      const hyperlink = schema.mark("hyperlink", {
        href: "https://example.test/field",
        _docxHyperlinkIndex: 1,
      });
      const pageBreak = schema.node("pageBreakRun", { clear: null });
      let children = [pageBreak, schema.node("textBoxAnchor", { anchorId: "field:0" })];
      if (unsupportedChild === "hyperlink") {
        children = [pageBreak.mark([hyperlink])];
      } else if (unsupportedChild === "bookmark") {
        children = [
          schema.node("bookmarkBoundary", { type: "start", id: 4, name: "field" }),
          pageBreak,
          schema.node("bookmarkBoundary", { type: "end", id: 4 }),
        ];
      }
      const structuredField = schema.node(
        "structuredField",
        {
          fieldType: "REF",
          instruction: "REF target",
          displayText: "",
          fieldKind: "complex",
        },
        children,
      );
      const prose = schema.node("doc", null, [schema.node("paragraph", null, [structuredField])]);

      const validation = validateProseMirrorDocument(prose);
      expect(validation.valid).toBe(false);
      let expectedMessage = "Complex field results cannot contain text-box anchors.";
      if (unsupportedChild === "hyperlink") {
        expectedMessage = "Complex field results cannot contain hyperlink content.";
      } else if (unsupportedChild === "bookmark") {
        expectedMessage = "Complex field results cannot contain bookmark boundaries.";
      }
      expect(validation.issues.map(({ message }) => message)).toContain(expectedMessage);
      expect(() => fromProseDoc(prose)).toThrow(
        "Cannot convert invalid ProseMirror document to DOCX model",
      );
    },
  );
});
