import { Result } from "better-result";
import { describe, expect, test } from "bun:test";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import { inspectDocxTemplate } from "../inspect";
import type { TemplatePlan, TemplatePlanItem } from "../types";
import { innerRangeOf, readRawPart } from "../xml";
import {
  buildMinimalDocx,
  buildTemplate,
  p,
  r,
  readPart,
  replacePart,
  TEMPLATE_IDS,
} from "./fixtures";

const unwrap = <T>(result: Result<T, unknown>): T => {
  if (Result.isError(result)) {
    throw result.error;
  }
  return result.value;
};

type PlanOptions = {
  source: Uint8Array;
  template: Uint8Array;
  items?: (ids: { id: string; hash: string }[]) => TemplatePlanItem[];
  extra?: Partial<TemplatePlan>;
};

const planFor = async ({
  source,
  template,
  items,
  extra = {},
}: PlanOptions): Promise<TemplatePlan> => {
  const manifest = unwrap(await analyzeDocxBlocks(source));
  const inspection = unwrap(await inspectDocxTemplate(template));
  const ids = manifest.blocks.map(({ id, hash }) => ({ id, hash }));
  return {
    version: 1,
    sourceSha256: manifest.sha256,
    templateSha256: inspection.sha256,
    body:
      items?.(ids) ??
      ids.map(({ id, hash }) => ({ kind: "source", blockId: id, hash })),
    omit: [],
    ...extra,
  };
};

describe("innerRangeOf", () => {
  test("finds the content after an opening tag whose attribute holds '>'", () => {
    const xml =
      '<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p w:x="a>b"><w:r/></w:p><w:q/></w:root>';
    const part = readRawPart("x", xml);
    const [first, second] = part?.root.elements ?? [];
    if (first === undefined || second === undefined) {
      throw new Error("fixture did not parse");
    }
    const inner = innerRangeOf(xml, first);

    expect(inner === null ? null : xml.slice(inner.start, inner.end)).toBe(
      "<w:r/>",
    );
    expect(innerRangeOf(xml, second)).toBeNull();
  });
});

describe("applyTemplatePlan edge content", () => {
  test("carries symbols, hyphens and breaks, and strips a label after leading space", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      [
        p(
          `${r("  (a)")}<w:r><w:tab/></w:r>${r("non")}<w:r><w:noBreakHyphen/></w:r>${r("breaking ")}<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r><w:r><w:br w:type="column"/></w:r>${r("after")}`,
        ),
      ].join(""),
    );
    const plan = await planFor({
      source,
      template,
      items: ([first]) =>
        first === undefined
          ? []
          : [
              {
                kind: "source",
                blockId: first.id,
                hash: first.hash,
                styleId: "HouseBody",
                stripLabel: "(a)",
              },
            ],
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";

    expect(report.blocks.at(0)).toMatchObject({
      text: "identical",
      strippedPrefix: "  (a)\t",
      sourceLabel: "(a)",
    });
    expect(unwrap(await analyzeDocxBlocks(docx)).blocks.at(0)?.text).toBe(
      "non‑breaking after",
    );
    expect(document).toContain("<w:noBreakHyphen/>");
    expect(document).toContain('<w:sym w:font="Wingdings" w:char="F0FC"/>');
    expect(document).toContain('<w:br w:type="column"/>');
  });

  test("drops a bookmark half whose other half sits in a dropped template block", async () => {
    const base = await buildTemplate();
    const xml = (await readPart(base, "word/document.xml")) ?? "";
    // The bookmark starts in the title and ends in the party paragraph.
    const template = await replacePart(
      base,
      "word/document.xml",
      xml
        .replace(
          '<w:t xml:space="preserve">Agreement</w:t>',
          '<w:t xml:space="preserve">Agreement</w:t></w:r><w:bookmarkStart w:id="40" w:name="Spanning"/><w:r>',
        )
        .replace(
          '<w:t xml:space="preserve">[party name]</w:t></w:r>',
          '<w:t xml:space="preserve">[party name]</w:t></w:r><w:bookmarkEnd w:id="40"/>',
        ),
    );
    const source = await buildMinimalDocx(p(r("Body text.")));
    const plan = await planFor({
      source,
      template,
      items: ([first]) => [
        { kind: "template", templateBlockIds: [TEMPLATE_IDS.title] },
        ...(first === undefined
          ? []
          : [
              {
                kind: "source" as const,
                blockId: first.id,
                hash: first.hash,
              },
            ]),
      ],
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";

    expect(report.ok).toBe(true);
    expect(document).toContain("Agreement");
    expect(document).not.toContain("Spanning");
    expect(document).not.toContain("bookmarkEnd");
  });

  test("patches a self-closed level and starts a list where the part holds none", async () => {
    const base = await buildTemplate();
    const numbering = (await readPart(base, "word/numbering.xml")) ?? "";
    const template = await replacePart(
      base,
      "word/numbering.xml",
      numbering.replace(
        /<w:lvl w:ilvl="4">.*?<\/w:lvl>/u,
        '<w:lvl w:ilvl="4"/>',
      ),
    );
    const source = await buildMinimalDocx(`${p(r("first"))}${p(r("second"))}`);
    const plan = await planFor({
      source,
      template,
      items: (ids) =>
        ids.map(({ id, hash }) => ({
          kind: "source",
          blockId: id,
          hash,
          styleId: "HouseBody",
          numbering: { numId: 1, ilvl: 4, restartFrom: 2, group: "romans" },
        })),
      extra: {
        numberingPatches: [
          {
            abstractNumId: 0,
            ilvl: 4,
            numFmt: "lowerRoman",
            lvlText: "(%5)",
            start: 1,
            suff: "space",
          },
        ],
      },
    });
    const { docx } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const patched = (await readPart(docx, "word/numbering.xml")) ?? "";

    expect(patched).toContain(
      '<w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:suff w:val="space"/><w:lvlText w:val="(%5)"/></w:lvl>',
    );
    expect(
      unwrap(await analyzeDocxBlocks(docx)).blocks.map(({ label }) => label),
    ).toEqual(["(ii)", "(iii)"]);
  });
});
