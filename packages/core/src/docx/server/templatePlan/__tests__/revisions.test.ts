import { beforeAll, describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import { inspectDocxTemplate } from "../inspect";
import { resolvePartRevisions } from "../revisions";
import type {
  DocxBlock,
  TemplatePlan,
  TemplatePlanTrackedView,
} from "../types";
import { buildMinimalDocx, buildTemplate, p, r, style } from "./fixtures";
import { codeOf, planFor, unwrap, verifyFor } from "./planning";

// Invented text; two authors and two dates.
const BY = 'w:author="First Author" w:date="2026-01-05T10:00:00Z"';
const BY_OTHER = 'w:author="Second Author" w:date="2026-02-07T09:30:00Z"';
const ins = (inner: string, id = 1): string =>
  `<w:ins w:id="${String(id)}" ${BY}>${inner}</w:ins>`;
const del = (text: string, id = 2): string =>
  `<w:del w:id="${String(id)}" ${BY_OTHER}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
/** A paragraph whose mark is tracked as inserted or deleted. */
const marked = (inner: string, kind: "ins" | "del", properties = ""): string =>
  p(inner, `${properties}<w:rPr><w:${kind} w:id="9" ${BY}/></w:rPr>`);

const texts = async (
  body: string,
  view: TemplatePlanTrackedView,
): Promise<string[]> => {
  const flat = (blocks: readonly DocxBlock[]): DocxBlock[] =>
    blocks.flatMap((block) => [block, ...flat(block.nested ?? [])]);
  const manifest = unwrap(
    await analyzeDocxBlocks(await buildMinimalDocx(body), {
      trackedChanges: view,
    }),
  );
  expect(manifest.trackedChanges).toBe(0);
  return flat(manifest.blocks)
    .filter(({ kind }) => kind === "paragraph")
    .map(({ text }) => text);
};

describe("a document with tracked changes, read in one view", () => {
  test("accepted, inserted text stays and deleted text goes; rejected, the reverse", async () => {
    const body = p(`${r("The fee is ")}${del("100")}${ins(r("120"))}${r(".")}`);

    expect(await texts(body, "accept")).toEqual(["The fee is 120."]);
    expect(await texts(body, "reject")).toEqual(["The fee is 100."]);
  });

  test("a bookmark that begins inside removed content keeps both its ends", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:del w:id="1" ${BY}><w:bookmarkStart w:id="5" w:name="kept"/><w:r><w:delText>gone</w:delText></w:r></w:del><w:r><w:t>stays</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p></w:body></w:document>`;

    expect(
      resolvePartRevisions("word/document.xml", xml, "accept")?.xml,
    ).toContain(
      '<w:p><w:bookmarkStart w:id="5" w:name="kept"/><w:r><w:t>stays</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p>',
    );
  });

  test("a moved passage stands where the view has it", async () => {
    const move = (kind: "moveFrom" | "moveTo", text: string) =>
      `<w:${kind} w:id="4" ${BY}>${kind === "moveFrom" ? `<w:r><w:delText>${text}</w:delText></w:r>` : r(text)}</w:${kind}>`;
    const body = [
      p(
        `<w:moveFromRangeStart w:id="3" w:name="move1" ${BY}/>${move("moveFrom", "Moved.")}<w:moveFromRangeEnd w:id="3"/>${r("First.")}`,
      ),
      p(
        `${r("Second.")}<w:moveToRangeStart w:id="5" w:name="move1" ${BY}/>${move("moveTo", "Moved.")}<w:moveToRangeEnd w:id="5"/>`,
      ),
    ].join("");

    expect(await texts(body, "accept")).toEqual(["First.", "Second.Moved."]);
    expect(await texts(body, "reject")).toEqual(["Moved.First.", "Second."]);
  });

  test("a paragraph mark the view removes joins its paragraph to the next, which keeps its own properties", async () => {
    const body = [
      marked(r("One "), "del", style("HouseTitle")),
      p(r("two."), style("HouseBody")),
      marked(r("Three "), "ins"),
      p(r("four.")),
    ].join("");
    const accepted = unwrap(
      await analyzeDocxBlocks(await buildMinimalDocx(body), {
        trackedChanges: "accept",
      }),
    );

    expect(accepted.blocks.map(({ text, styleId }) => [text, styleId])).toEqual(
      [
        ["One two.", "HouseBody"],
        ["Three ", null],
        ["four.", null],
      ],
    );
    expect(accepted.trackedChangesResolved).toMatchObject({
      view: "accept",
      paragraphMarks: 1,
      authors: ["First Author"],
    });
    expect(await texts(body, "reject")).toEqual([
      "One ",
      "two.",
      "Three four.",
    ]);
    // A mark that goes with nothing after it to join stays a mark.
    expect(await texts(marked(r("Last."), "del"), "accept")).toEqual(["Last."]);
  });

  test("a table row the view removes goes whole", async () => {
    const row = (text: string, tracked = "") =>
      `<w:tr>${tracked === "" ? "" : `<w:trPr><w:${tracked} w:id="7" ${BY}/></w:trPr>`}<w:tc>${p(r(text))}</w:tc></w:tr>`;
    const body = `<w:tbl>${row("kept")}${row("added", "ins")}${row("struck", "del")}</w:tbl>${p(r("after"))}`;

    expect(await texts(body, "accept")).toEqual(["kept", "added", "after"]);
    expect(await texts(body, "reject")).toEqual(["kept", "struck", "after"]);
  });

  test("rejected, recorded formatting comes back; accepted, it is dropped", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="New"/><w:rPr><w:b/></w:rPr><w:pPrChange w:id="1" ${BY}><w:pPr><w:pStyle w:val="Old"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:i/><w:rPrChange w:id="2" ${BY_OTHER}><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>text</w:t></w:r></w:p></w:body></w:document>`;
    const accepted = resolvePartRevisions("word/document.xml", xml, "accept");
    const rejected = resolvePartRevisions("word/document.xml", xml, "reject");

    expect(accepted?.xml).toContain(
      '<w:pPr><w:pStyle w:val="New"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>text</w:t></w:r>',
    );
    // The mark's own run properties stay with the paragraph.
    expect(rejected?.xml).toContain(
      '<w:pPr><w:pStyle w:val="Old"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>text</w:t></w:r>',
    );
    expect(rejected?.xml).not.toContain("Change");
    expect(accepted?.tally.formattingChanges).toBe(2);
    expect([...(accepted?.tally.authors ?? [])].toSorted()).toEqual([
      "First Author",
      "Second Author",
    ]);
  });

  test("a document without tracked changes is read as it is", async () => {
    const plain = await buildMinimalDocx(p(r("Plain.")));
    const read = unwrap(
      await analyzeDocxBlocks(plain, { trackedChanges: "reject" }),
    );
    const asIs = unwrap(await analyzeDocxBlocks(plain));

    expect(read.blocks).toEqual(asIs.blocks);
    expect(read.sha256).toBe(asIs.sha256);
  });
});

describe("a plan for one view of tracked changes", () => {
  let template: Uint8Array;
  beforeAll(async () => {
    template = await buildTemplate();
  });
  const tracked = async (): Promise<Uint8Array> =>
    await buildMinimalDocx(
      [
        p(`${r("The fee is ")}${del("100")}${ins(r("120"))}${r(".")}`),
        marked(r("It is due "), "del"),
        p(r("on delivery.")),
      ].join(""),
    );

  test("the manifest of a tracked document says what it tracks and keeps the file's hash in a view", async () => {
    const source = await tracked();
    const asIs = unwrap(await analyzeDocxBlocks(source));
    const accepted = unwrap(
      await analyzeDocxBlocks(source, { trackedChanges: "accept" }),
    );

    expect(asIs.trackedChanges).toBeGreaterThan(0);
    expect(asIs.revisions).toEqual({
      insertions: 1,
      deletions: 2,
      moves: 0,
      paragraphMarks: 1,
      tableRows: 0,
      formattingChanges: 0,
      authors: ["First Author", "Second Author"],
      first: "2026-01-05T10:00:00Z",
      last: "2026-02-07T09:30:00Z",
    });
    expect(accepted.sha256).toBe(asIs.sha256);
    expect(accepted.revisions).toBeUndefined();
  });

  test("applies in the chosen view, writes no revision and reports what was resolved", async () => {
    const source = await tracked();
    for (const [view, expected] of [
      ["accept", ["The fee is 120.", "It is due on delivery."]],
      ["reject", ["The fee is 100.", "It is due ", "on delivery."]],
    ] as const) {
      const plan = await planFor({
        source,
        template,
        items: (blocks) =>
          blocks.map(({ id, hash }) => ({
            kind: "source",
            blockId: id,
            hash,
            styleId: "HouseBody",
          })),
        extra: { trackedChanges: view },
      });
      const { docx, report } = unwrap(
        await applyTemplatePlan({ source, template, plan }),
      );
      const output = unwrap(await analyzeDocxBlocks(docx));
      const document =
        (await (
          await ZipArchive.read(docx)
        )
          .entry("word/document.xml")
          ?.text()) ?? "";

      expect(report.ok).toBe(true);
      expect(report.trackedChanges).toMatchObject({ view, insertions: 1 });
      expect(output.trackedChanges).toBe(0);
      expect(document).not.toMatch(/<w:(?:ins|del|delText)\b/u);
      expect(
        output.blocks
          .map(({ text }) => text)
          .filter((text) => expected.some((line) => line === text)),
      ).toEqual([...expected]);
      // Verification holds the output to the view the plan names.
      const other = view === "accept" ? "reject" : "accept";
      expect(
        codeOf(
          await applyTemplatePlan({
            source,
            template,
            plan: { ...plan, trackedChanges: other },
          }),
        ),
      ).not.toBe("applied");
      expect(
        (await verifyFor({ source, template, output: docx, plan })).ok,
      ).toBe(true);
    }
  });

  test("without a view the source is still refused, and an unknown view is no plan", async () => {
    const source = await tracked();
    const plan = await planFor({ source, template });

    expect(codeOf(await applyTemplatePlan({ source, template, plan }))).toBe(
      "tracked-changes",
    );
    // A plan read from JSON may name anything.
    const unknownView: TemplatePlan = JSON.parse(
      JSON.stringify({ ...plan, trackedChanges: "merge" }),
    );
    expect(
      codeOf(await applyTemplatePlan({ source, template, plan: unknownView })),
    ).toBe("plan-invalid");
  });
});

describe("a plan that keeps the tracked changes of its source", () => {
  let template: Uint8Array;
  beforeAll(async () => {
    template = await buildTemplate();
  });
  const partOf = async (docx: Uint8Array, path: string): Promise<string> =>
    (await (await ZipArchive.read(docx)).entry(path)?.text()) ?? "";
  /** A plan made on the accepted view's blocks, as a caller makes it. */
  const keeping = async (
    source: Uint8Array,
    items?: Parameters<typeof planFor>[0]["items"],
  ): Promise<TemplatePlan> => ({
    ...(await planFor({
      source,
      template,
      ...(items === undefined ? {} : { items }),
      extra: { trackedChanges: "accept" },
    })),
    trackedChanges: "keep",
  });
  const body = (blocks: { id: string; hash: string }[]) =>
    blocks.map(({ id, hash }) => ({
      kind: "source" as const,
      blockId: id,
      hash,
      styleId: "HouseBody",
    }));

  test("the output shows the same insertions and deletions, by the same authors, and tracks nothing new", async () => {
    const source = await buildMinimalDocx(
      `${p(`${r("The fee is ")}${del("100")}${ins(r("120"))}${r(".")}`)}${p(r("Untouched."))}`,
    );
    const plan = await keeping(source, body);
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = await partOf(docx, "word/document.xml");

    expect(report.ok).toBe(true);
    expect(report.trackedChangesKept).toMatchObject({
      insertions: 1,
      deletions: 1,
      authors: ["First Author", "Second Author"],
    });
    expect(document).toMatch(
      /<w:del w:id="\d+" w:author="Second Author" w:date="2026-02-07T09:30:00Z"><w:r><w:delText xml:space="preserve">100<\/w:delText><\/w:r><\/w:del><w:ins w:id="\d+" w:author="First Author" w:date="2026-01-05T10:00:00Z"><w:r><w:t xml:space="preserve">120<\/w:t><\/w:r><\/w:ins>/u,
    );
    // Both readings of the output are the source's.
    for (const [view, text] of [
      ["accept", "The fee is 120."],
      ["reject", "The fee is 100."],
    ] as const) {
      expect(
        unwrap(await analyzeDocxBlocks(docx, { trackedChanges: view }))
          .blocks.map((block) => block.text)
          .filter((line) => line.startsWith("The fee")),
      ).toEqual([text]);
    }
    expect(await partOf(docx, "word/settings.xml")).not.toContain(
      "trackRevisions",
    );
    // Also on a template that has tracking switched on.
    const settings = await partOf(template, "word/settings.xml");
    const tracking = await ZipArchive.read(template);
    tracking.set(
      "word/settings.xml",
      settings.replace(/(<w:settings\b[^>]*>)/u, "$1<w:trackRevisions/>"),
    );
    const onTracking = await tracking.generate();
    expect(await partOf(onTracking, "word/settings.xml")).toContain(
      "<w:trackRevisions/>",
    );
    const switchedOff = unwrap(
      await applyTemplatePlan({
        source,
        template: onTracking,
        plan: {
          ...plan,
          templateSha256: unwrap(await inspectDocxTemplate(onTracking)).sha256,
        },
      }),
    );
    expect(await partOf(switchedOff.docx, "word/settings.xml")).not.toContain(
      "trackRevisions",
    );
    // Verification holds the output to each change and to who made it.
    const tampered = [
      document.replace('w:author="Second Author"', 'w:author="Someone"'),
      document.replace(/<w:del [^>]*>.*?<\/w:del>/u, ""),
      document.replace(">120<", ">125<"),
    ];
    for (const xml of tampered) {
      expect(xml).not.toBe(document);
      const zip = await ZipArchive.read(docx);
      zip.set("word/document.xml", xml);
      expect(
        (
          await verifyFor({
            source,
            template,
            output: await zip.generate(),
            plan,
          })
        ).issues.map(({ code }) => code),
      ).toEqual(["text-mismatch"]);
    }
  });

  test("a change that the two views read as other paragraphs is refused by name", async () => {
    const cases: [string, string][] = [
      [
        `${marked(r("One "), "del")}${p(r("two."))}`,
        "inserted or deleted paragraph mark",
      ],
      [
        p(
          `${r("First.")}<w:moveFrom w:id="4" ${BY}><w:r><w:delText>Moved.</w:delText></w:r></w:moveFrom>`,
        ),
        "move",
      ],
      [
        p(`${r("An ")}${ins(`${r("added ")}${del("old ", 7)}`)}${r("word.")}`),
        "inside another change",
      ],
    ];
    for (const [sourceBody, named] of cases) {
      const source = await buildMinimalDocx(sourceBody);
      const applied = await applyTemplatePlan({
        source,
        template,
        plan: await keeping(source, body),
      });
      expect(codeOf(applied)).toBe("tracked-changes");
      expect(applied.isErr() && applied.error.message).toContain(named);
      // The same answer to a caller who asks before planning.
      const asked = await analyzeDocxBlocks(source, { trackedChanges: "keep" });
      expect(asked.isErr() && asked.error.code).toBe("tracked-changes");
    }
  });

  test("a label that holds a change is not stripped", async () => {
    const source = await buildMinimalDocx(
      p(`${ins(r("3."))}<w:r><w:tab/></w:r>${r("The term.")}`),
    );
    const numbered = await keeping(source, ([first]) =>
      first === undefined
        ? []
        : [
            {
              kind: "source",
              blockId: first.id,
              hash: first.hash,
              styleId: "HouseBody",
              stripLabel: "3.",
              numbering: { numId: 1, ilvl: 1 },
            },
          ],
    );
    expect(
      codeOf(await applyTemplatePlan({ source, template, plan: numbered })),
    ).toBe("strip-label-mismatch");
    // Placed with its typed number, the change goes with it.
    const { docx, report } = unwrap(
      await applyTemplatePlan({
        source,
        template,
        plan: await keeping(source, body),
      }),
    );
    expect(report.ok).toBe(true);
    expect(await partOf(docx, "word/document.xml")).toMatch(
      /<w:ins [^>]*><w:r><w:t xml:space="preserve">3\.<\/w:t><\/w:r><\/w:ins>/u,
    );
  });
});
