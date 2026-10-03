import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  SPLIT_HALVES,
  type DocumentOp,
} from "@stll/docx-core/ops";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Paragraph, Run } from "../types/document";
import { parseDocumentBody } from "./documentParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const stamp = { id: 1000, author: "Joining", date: "2026-02-03T04:05:06Z" };

const BOUNDARIES = {
  plain: "",
  bookmark: '<w:bookmarkStart w:id="20" w:name="seam"/><w:bookmarkEnd w:id="20"/>',
  comment: '<w:commentRangeStart w:id="20"/><w:commentRangeEnd w:id="20"/>',
  field: '<w:fldSimple w:instr="DATE"><w:r><w:t>cached</w:t></w:r></w:fldSimple>',
  revision:
    '<w:ins w:id="20" w:author="Earlier" w:date="2026-01-01T00:00:00Z"><w:r><w:t>pending</w:t></w:r></w:ins>',
  control: '<w:sdt><w:sdtPr><w:id w:val="20"/><w:richText/></w:sdtPr><w:sdtContent/></w:sdt>',
  tab: "<w:r><w:tab/></w:r>",
  break: "<w:r><w:br/></w:r>",
  propertyRevision:
    '<w:r><w:rPr><w:b/><w:rPrChange w:id="20" w:author="Earlier" w:date="2026-01-01T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>reviewed</w:t></w:r>',
} as const;

const wrapBody = (paragraphs: string) =>
  `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphs}</w:body></w:document>`;

const paragraphsOf = (document: Document): Paragraph[] =>
  document.package.document.content.map((block) => {
    if (block.type !== "paragraph") panic("The generated body contains only paragraphs");
    return block;
  });

const serialize = (document: Document) => paragraphsOf(document).map(serializeParagraph).join("");

const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value;
};

const assertJoinLaws = (document: Document) => {
  const join = {
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    nextBlockId: "00000002",
    survivor: SPLIT_HALVES.SECOND,
    depth: 0,
  } as const satisfies DocumentOp;
  const direct = apply(document, [join]);
  const tracked = apply(document, [{ ...join, revision: stamp, newIds: { revision: [1001] } }]);
  const accepted = apply(tracked.document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.revisions,
      decision: REVISION_DECISIONS.ACCEPT,
    },
  ]);
  expect(accepted.document).toStrictEqual(direct.document);
  expect(apply(direct.document, direct.inverse).document).toStrictEqual(document);
  expect(apply(tracked.document, tracked.inverse).document).toStrictEqual(document);
  expect(apply(accepted.document, accepted.inverse).document).toStrictEqual(tracked.document);
  return direct.document;
};

const assertSavedJoinLaws = (document: Document) => {
  const joined = assertJoinLaws(document);
  const saved = serialize(joined);
  const reopened = { package: { document: parseDocumentBody(wrapBody(saved)) } };
  expect(serialize(reopened)).toBe(saved);
  expect(paragraphsOf(reopened)).toStrictEqual(paragraphsOf(joined));
  expect(paragraphsOf(reopened).map(({ paraId }) => paraId)).toStrictEqual(["00000002"]);
  return joined;
};

const text = fc
  .array(fc.constantFrom("a", "b", "é", "😀", " "), { minLength: 1, maxLength: 8 })
  .map((letters) => letters.join(""));

test("joined seams reach the parser's fixed point without crossing authored boundaries", () => {
  assertProperty(
    fc.property(text, text, (left, right) => {
      for (const italic of [false, true]) {
        const properties = `<w:rPr><w:b/>${italic ? "<w:i/>" : ""}</w:rPr>`;
        const run = (value: string) =>
          `<w:r>${properties}<w:t xml:space="preserve">${value}</w:t></w:r>`;
        for (const [boundary, markup] of Object.entries(BOUNDARIES)) {
          const document = {
            package: {
              document: parseDocumentBody(
                wrapBody(
                  `<w:p w14:paraId="00000001">${run(left)}${markup}</w:p>` +
                    `<w:p w14:paraId="00000002">${run(right)}</w:p>`,
                ),
              ),
            },
          };
          const joined = paragraphsOf(assertSavedJoinLaws(document)).at(0);
          if (joined === undefined) panic("A join leaves its surviving paragraph");
          if (boundary === "plain") {
            expect(joined.content).toHaveLength(1);
            expect(joined.content.at(0)?.type).toBe("run");
            continue;
          }
          // No item separates these runs after the join except this authored boundary.
          expect(joined.content).toStrictEqual(
            paragraphsOf(document).flatMap(({ content }) => content),
          );
        }
      }
    }),
    { numRuns: 40 },
  );
});

test("different run records keep both sides of a joined seam", () => {
  assertProperty(
    fc.property(text, text, (left, right) => {
      for (const [leftProps, rightProps, leftAttrs, rightAttrs] of [
        ["<w:b/>", "<w:i/>", "", ""],
        ["<w:b/>", "<w:b/>", ' w:rsidR="00112233"', ' w:rsidR="00445566"'],
        ["<w:b/><w:webHidden/>", "<w:b/>", "", ""],
      ] as const) {
        const document = {
          package: {
            document: parseDocumentBody(
              wrapBody(
                `<w:p w14:paraId="00000001"><w:r${leftAttrs}><w:rPr>${leftProps}</w:rPr><w:t xml:space="preserve">${left}</w:t></w:r></w:p>` +
                  `<w:p w14:paraId="00000002"><w:r${rightAttrs}><w:rPr>${rightProps}</w:rPr><w:t xml:space="preserve">${right}</w:t></w:r></w:p>`,
              ),
            ),
          },
        };
        const joined = paragraphsOf(assertSavedJoinLaws(document)).at(0);
        if (joined === undefined) panic("A join leaves its surviving paragraph");
        expect(joined.content).toStrictEqual(
          paragraphsOf(document).flatMap(({ content }) => content),
        );
      }
    }),
    { numRuns: 25 },
  );
});

test("alike authored containers on opposite sides of a join retain their identities", () => {
  assertProperty(
    fc.property(text, text, (left, right) => {
      const run = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
      const containers = {
        revision: (value: string, id: number) =>
          `<w:ins w:id="${id}" w:author="Earlier" w:date="2026-01-01T00:00:00Z">${run(value)}</w:ins>`,
        control: (value: string, id: number) =>
          `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:richText/></w:sdtPr><w:sdtContent>${run(value)}</w:sdtContent></w:sdt>`,
        hyperlink: (value: string) => `<w:hyperlink w:anchor="seam">${run(value)}</w:hyperlink>`,
        field: (value: string) => `<w:fldSimple w:instr="DATE">${run(value)}</w:fldSimple>`,
      };
      for (const container of Object.values(containers)) {
        const document = {
          package: {
            document: parseDocumentBody(
              wrapBody(
                `<w:p w14:paraId="00000001">${container(left, 20)}</w:p>` +
                  `<w:p w14:paraId="00000002">${container(right, 21)}</w:p>`,
              ),
            ),
          },
        };
        const joined = paragraphsOf(assertSavedJoinLaws(document)).at(0);
        if (joined === undefined) panic("A join leaves its surviving paragraph");
        expect(joined.content).toStrictEqual(
          paragraphsOf(document).flatMap(({ content }) => content),
        );
      }
    }),
    { numRuns: 25 },
  );
});

test("a join changes only its seam even when other authored runs could merge", () => {
  assertProperty(
    fc.property(text, text, (left, right) => {
      const run = (value: string) =>
        ({ type: "run", content: [{ type: "text", text: value }] }) satisfies Run;
      const document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "00000001",
                content: [run("prefix"), run("left"), run(left)],
              },
              {
                type: "paragraph",
                paraId: "00000002",
                content: [run(right), run("right"), run("suffix")],
              },
            ],
          },
        },
      } as const satisfies Document;
      const joined = paragraphsOf(assertJoinLaws(document)).at(0);
      if (joined === undefined) panic("A join leaves its surviving paragraph");
      expect(joined.content).toStrictEqual([
        run("prefix"),
        run("left"),
        run(left + right),
        run("right"),
        run("suffix"),
      ]);
    }),
    { numRuns: 25 },
  );
});

test("equal run attributes in different orders merge and restore each original order", () => {
  assertProperty(
    fc.property(text, text, (left, right) => {
      const document = {
        package: {
          document: parseDocumentBody(
            wrapBody(
              `<w:p w14:paraId="00000001"><w:r w:rsidR="00112233" w:rsidRPr="00445566"><w:t xml:space="preserve">${left}</w:t></w:r></w:p>` +
                `<w:p w14:paraId="00000002"><w:r w:rsidRPr="00445566" w:rsidR="00112233"><w:t xml:space="preserve">${right}</w:t></w:r></w:p>`,
            ),
          ),
        },
      };
      const joined = paragraphsOf(assertSavedJoinLaws(document)).at(0);
      if (joined === undefined) panic("A join leaves its surviving paragraph");
      expect(joined.content).toHaveLength(1);
    }),
    { numRuns: 25 },
  );
});

test("equivalent nested formatting spellings merge and undo restores their exact records", () => {
  assertProperty(
    fc.property(text, text, (left, right) => {
      const body = parseDocumentBody(
        wrapBody(
          `<w:p w14:paraId="00000001"><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${left}</w:t></w:r></w:p>` +
            `<w:p w14:paraId="00000002"><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${right}</w:t></w:r></w:p>`,
        ),
      );
      // Use the operation owner so the parsed body's section projection stays consistent.
      const document = apply({ package: { document: body } }, [
        {
          type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
          from: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 },
          to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: right.length },
          patch: { language: {} },
        },
      ]).document;
      const joined = paragraphsOf(assertSavedJoinLaws(document)).at(0);
      if (joined === undefined) panic("A join leaves its surviving paragraph");
      expect(joined.content).toHaveLength(1);
    }),
    { numRuns: 25 },
  );
});

test("splitting authored alike runs restores their segmentation and identities exactly", () => {
  assertProperty(
    fc.property(text, text, (left, right) => {
      const run = (value: string) =>
        ({ type: "run", content: [{ type: "text", text: value }] }) satisfies Run;
      const document = {
        package: {
          document: {
            content: [{ type: "paragraph", paraId: "00000001", content: [run(left), run(right)] }],
          },
        },
      } as const satisfies Document;
      for (const newHalf of Object.values(SPLIT_HALVES)) {
        for (const offset of [0, left.length, left.length + right.length]) {
          const split = apply(document, [
            {
              type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
              at: { story: OP_STORIES.MAIN, blockId: "00000001", offset },
              newBlockId: "00000002",
              newHalf,
            },
          ]);
          expect(apply(split.document, split.inverse).document).toStrictEqual(document);
        }
      }
    }),
    { numRuns: 25 },
  );
});
