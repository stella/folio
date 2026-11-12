/**
 * Formatting part of a run that also holds a line break or a tab cuts the run
 * into pieces, and a save writes each piece as its own `w:r`. A piece of text
 * alone reads back without the identity that grouped the leaves of the
 * authored run, so its neighbours could no longer take it back: resolving the
 * change rejoined one run in the editor that made it and left three after a
 * save and reopen. Resolving gives the same paragraph either way, and
 * rejecting gives back the authored run.
 */

import { describe, expect, test } from "bun:test";

import { canonicalResolvedReviewModel } from "../../../../test/consumer-scenarios/support/review-comparison";
import { createDocx } from "../docx/rezip";
import { consolidateParagraphContent } from "../docx/runConsolidator";
import type {
  Document,
  ParagraphContent,
  Run,
  RunContent,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIEditApplyMode } from "./types";

const AUTHOR = "Reviewer";
const LINES = ["Delivery schedule", "Payment terms", "Notice period"] as const;

const SEPARATORS = [
  { name: "a line break", content: { type: "break" }, text: "\n" },
  { name: "a tab", content: { type: "tab" }, text: "\t" },
] as const satisfies readonly {
  name: string;
  content: RunContent;
  text: string;
}[];
type Separator = (typeof SEPARATORS)[number];

const HOSTS = ["paragraph", "hyperlink"] as const;
type Host = (typeof HOSTS)[number];

/** Where the formatted text sits against the separators, by block text. */
const RANGES = [
  { name: "first line start", from: "Delivery", length: "Delivery".length },
  { name: "middle line", from: "Payment", length: "Payment".length },
  { name: "last line end", from: "period", length: "period".length },
  {
    name: "across a separator",
    from: "schedule",
    length: "schedule".length + 1 + "Payment".length,
  },
] as const;
type Range = (typeof RANGES)[number];

const MODES = [
  "direct",
  "tracked-changes",
  "suggested",
] as const satisfies readonly FolioAIEditApplyMode[];

const authoredRun = (separator: Separator): Run => ({
  type: "run",
  content: LINES.flatMap((text, index): RunContent[] =>
    index === 0
      ? [{ type: "text", text }]
      : [{ ...separator.content }, { type: "text", text }],
  ),
});

/** What rejecting gives back: the authored run, whole and unformatted. */
const authoredRuns = (separator: Separator) => [
  {
    content: authoredRun(separator).content.map((item) =>
      item.type === "text" ? item.text : item.type,
    ),
  },
];

const baseline = (host: Host, separator: Separator): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  const run = authoredRun(separator);
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "1B0E5A11",
      content: [
        host === "paragraph"
          ? run
          : {
              type: "hyperlink",
              href: "https://example.com/terms",
              children: [run],
            },
      ],
    },
  ];
  return createDocx(document);
};

const runsOf = (content: readonly ParagraphContent[]): Run[] =>
  content.flatMap((item) => {
    if (item.type === "run") return [item];
    if (item.type === "hyperlink") {
      return item.children.filter((child) => child.type === "run");
    }
    return [];
  });

/** The paragraph's runs after the comparison's consolidation, by content and formatting. */
const paragraphRuns = (document: Document) => {
  const paragraph = document.package.document.content.find(
    (block) => block.type === "paragraph",
  );
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected the authored paragraph");
  }
  return runsOf(consolidateParagraphContent(paragraph.content)).map(
    ({ content, formatting }) => {
      const shape: { content: string[]; formatting?: Run["formatting"] } = {
        content: content.map((item) =>
          item.type === "text" ? item.text : item.type,
        ),
      };
      if (formatting && Object.keys(formatting).length > 0) {
        shape.formatting = formatting;
      }
      return shape;
    },
  );
};

type ApplyOptions = {
  host: Host;
  separator: Separator;
  range: Range;
  mode: FolioAIEditApplyMode;
};

const formatted = async ({ host, separator, range, mode }: ApplyOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await baseline(host, separator),
    { author: AUTHOR },
  );
  const block = reviewer.snapshot().blocks.at(0);
  if (!block) throw new Error("Expected the authored block");
  expect(block.text).toBe(LINES.join(separator.text));
  const startOffset = block.text.indexOf(range.from);
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode,
      operations: [
        {
          id: "underline",
          type: "formatRange",
          range: createFolioAITextRangeHandle({
            blockId: block.id,
            text: block.text,
            startOffset,
            endOffset: startOffset + range.length,
          }),
          formatting: { underline: true },
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(1);
  return { reviewer, suggestionId: result.applied.at(0)?.suggestionId };
};

const reopen = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: AUTHOR });

const cases = HOSTS.flatMap((host) =>
  SEPARATORS.flatMap((separator) =>
    RANGES.map((range) => ({ host, separator, range })),
  ),
);

describe("formatting text beside a line break or a tab", () => {
  for (const { host, separator, range } of cases) {
    const label = `${host}, ${separator.name}, ${range.name}`;

    test(`${label}: direct formatting reads the same after save`, async () => {
      const { reviewer } = await formatted({
        host,
        separator,
        range,
        mode: "direct",
      });
      const reopened = await reopen(reviewer);

      expect(paragraphRuns(reopened.toDocument())).toEqual(
        paragraphRuns(reviewer.toDocument()),
      );
    });

    for (const mode of MODES.filter((candidate) => candidate !== "direct")) {
      test(`${label}: ${mode} formatting resolves the same with and without a save`, async () => {
        const { reviewer, suggestionId } = await formatted({
          host,
          separator,
          range,
          mode,
        });
        if (mode === "suggested") {
          if (suggestionId === undefined) {
            throw new Error("Expected a suggestion id");
          }
          expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
        }
        const reopened = await reopen(reviewer);
        const direct = await formatted({
          host,
          separator,
          range,
          mode: "direct",
        });

        for (const decision of ["accept", "reject"] as const) {
          const live = reviewer.toResolvedDocument(decision);
          const saved = reopened.toResolvedDocument(decision);
          expect(canonicalResolvedReviewModel(saved)).toEqual(
            canonicalResolvedReviewModel(live),
          );
          expect(paragraphRuns(saved)).toEqual(paragraphRuns(live));
        }
        expect(paragraphRuns(reopened.toResolvedDocument("reject"))).toEqual(
          authoredRuns(separator),
        );
        expect(paragraphRuns(reopened.toResolvedDocument("accept"))).toEqual(
          paragraphRuns(direct.reviewer.toDocument()),
        );
      });
    }

    test(`${label}: a rejected suggestion leaves the authored run`, async () => {
      const { reviewer, suggestionId } = await formatted({
        host,
        separator,
        range,
        mode: "suggested",
      });
      if (suggestionId === undefined) {
        throw new Error("Expected a suggestion id");
      }
      expect(reviewer.rejectSuggestion(suggestionId)).toBe(true);
      const reopened = await reopen(reviewer);

      for (const document of [reviewer.toDocument(), reopened.toDocument()]) {
        expect(paragraphRuns(document)).toEqual(authoredRuns(separator));
      }
    });
  }
});
