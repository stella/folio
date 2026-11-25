import { Result } from "better-result";
import { describe, expect, test } from "bun:test";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import { createRelationshipsEditor } from "../emit";
import { inspectDocxTemplate } from "../inspect";
import { sha256Hex } from "../package";
import { objectFingerprint } from "../resources";
import { readRawPart } from "../xml";
import {
  buildMinimalDocx,
  buildTemplate,
  fieldRuns,
  NS,
  p,
  r,
  readPart,
  replacePart,
  XML,
} from "./fixtures";
import {
  applyFor,
  buildSourceWith,
  codeOf,
  issuesOf,
  planFor,
  relationshipsXml,
  revision,
  unwrap,
  verifyFor,
} from "./planning";

const MATH =
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const NOTE_MARK = "<w:r><w:footnoteRef/></w:r>";

const note = (id: number, inner: string): string =>
  `<w:footnote w:id="${id}">${p(`${NOTE_MARK}${inner}`)}</w:footnote>`;

const noteReference = (id: number): string =>
  `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;

const OPEN_FIELD =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';

const errorOf = <T>(result: Result<T, { code: string; part?: string }>) =>
  Result.isError(result)
    ? { code: result.error.code, part: result.error.part }
    : null;

describe("stories with a field left open", () => {
  test("a note keeps its text after a note whose field never closes", async () => {
    const source = await buildSourceWith({
      body: p(`${r("Text")}${noteReference(1)}${noteReference(2)}`),
      footnotes: `${note(1, `${OPEN_FIELD}${r("cached")}`)}${note(2, r(" Plain note."))}`,
    });
    const manifest = unwrap(await analyzeDocxBlocks(source));

    expect(manifest.footnotes.at(1)).toEqual({ id: "2", text: " Plain note." });
    // The unpaired mark still refuses the source as a whole.
    expect(
      codeOf(await applyFor({ source, template: await buildTemplate() })),
    ).toBe("unsupported-content");
  });

  test("a template note's field result stops at the end of its note", async () => {
    const base = await buildTemplate();
    const template = await replacePart(
      base,
      "word/footnotes.xml",
      `${XML}<w:footnotes ${NS}>${note(1, `${OPEN_FIELD}${r("cached")}`)}${note(2, r(" Plain note."))}</w:footnotes>`,
    );
    const inspection = unwrap(await inspectDocxTemplate(template));

    expect(
      inspection.fields
        .filter(({ part }) => part === "word/footnotes.xml")
        .map(({ kind, result }) => ({ kind, result })),
    ).toEqual([{ kind: "DATE", result: "cached" }]);
  });
});

describe("parts the package cannot give", () => {
  test("a header without readable XML refuses the inspection and the apply", async () => {
    const template = await replacePart(
      await buildTemplate(),
      "word/header1.xml",
      "not a part",
    );
    const source = await buildMinimalDocx(p(r("Body text.")));

    expect(errorOf(await inspectDocxTemplate(template))).toEqual({
      code: "xml-unreadable",
      part: "word/header1.xml",
    });
    const manifest = unwrap(await analyzeDocxBlocks(source));
    const applied = await applyTemplatePlan({
      source,
      template,
      plan: {
        version: 1,
        sourceSha256: manifest.sha256,
        templateSha256: sha256Hex(template),
        body: manifest.blocks.map(({ id, hash }) => ({
          kind: "source",
          blockId: id,
          hash,
        })),
        omit: [],
      },
    });
    expect(errorOf(applied)).toEqual({
      code: "xml-unreadable",
      part: "word/header1.xml",
    });
    // A template that names no header still inspects.
    expect(
      Result.isOk(await inspectDocxTemplate(await buildMinimalDocx(p(r("x"))))),
    ).toBe(true);
  });

  test("a part the document names and the package lacks is refused", async () => {
    const docx = await buildMinimalDocx(p(r("Body text.")), {
      "word/_rels/document.xml.rels": relationshipsXml([
        { id: "rId1", type: "styles", target: "styles.xml" },
      ]),
    });

    expect(errorOf(await analyzeDocxBlocks(docx))).toEqual({
      code: "part-missing",
      part: "word/styles.xml",
    });
  });

  test("new notes are refused where another part has their name", async () => {
    // The template has no notes; its styles live under the usual notes name.
    const base = await buildTemplate({ withoutFootnotes: true });
    const relationships =
      (await readPart(base, "word/_rels/document.xml.rels")) ?? "";
    const styles = (await readPart(base, "word/styles.xml")) ?? "";
    const template = await replacePart(
      await replacePart(base, "word/footnotes.xml", styles),
      "word/_rels/document.xml.rels",
      relationships.replace('Target="styles.xml"', 'Target="footnotes.xml"'),
    );
    const withNote = await buildSourceWith({
      body: p(`${r("Text")}${noteReference(1)}`),
      footnotes: note(1, r(" A note.")),
    });
    const refused = await applyFor({ source: withNote, template });

    expect(errorOf(refused)).toEqual({
      code: "unsupported-content",
      part: "word/footnotes.xml",
    });
    // Without notes to write, the part is left alone and the apply goes on.
    const { docx } = unwrap(
      await applyFor({
        source: await buildMinimalDocx(p(r("Body text."))),
        template,
      }),
    );
    expect(await readPart(docx, "word/footnotes.xml")).toBe(styles);
  });
});

describe("tracked content", () => {
  test("counts a revision inside an equation", async () => {
    const equation = (inner: string): string =>
      `<w:p ${MATH}><m:oMath>${inner}</m:oMath></w:p>`;
    const run = "<m:r><m:t>x</m:t></m:r>";
    const tracked = await buildMinimalDocx(
      equation(`<w:ins ${revision(1)}>${run}</w:ins>`),
    );
    const plain = await buildMinimalDocx(equation(run));

    expect(unwrap(await analyzeDocxBlocks(tracked)).trackedChanges).toBe(1);
    expect(unwrap(await analyzeDocxBlocks(plain)).trackedChanges).toBe(0);
  });

  test("reads the blocks of an insertion and leaves out those of a deletion", async () => {
    const docx = await buildMinimalDocx(
      `<w:ins ${revision(1)}>${p(r("Inserted."))}</w:ins><w:del ${revision(2)}>${p(r("Deleted."))}</w:del>${p(r("Next."))}`,
    );
    const manifest = unwrap(await analyzeDocxBlocks(docx));

    expect(manifest.blocks.map(({ id, text }) => ({ id, text }))).toEqual([
      { id: "p0", text: "Inserted." },
      { id: "p1", text: "Next." },
    ]);
    expect(manifest.trackedChanges).toBe(2);
  });
});

describe("ids and fingerprints", () => {
  test("relationship ids stay distinct past a suffix too long to count", () => {
    const editor = createRelationshipsEditor(
      "word/_rels/document.xml.rels",
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId99999999999999999999" Type="urn:t" Target="a.xml"/><Relationship Id="rId1" Type="urn:t" Target="b.xml"/></Relationships>`,
    );
    const first = editor?.add({
      type: "urn:t",
      target: "c.xml",
      external: false,
    });
    const second = editor?.add({
      type: "urn:t",
      target: "d.xml",
      external: false,
    });
    const again = editor?.add({
      type: "urn:t",
      target: "c.xml",
      external: false,
    });

    expect(first).toBe("rId2");
    expect(second).toBe("rId3");
    expect(again).toBe(first);
  });

  test("an attribute value cannot pass for a second attribute", () => {
    const fingerprint = (attributes: string): string => {
      const part = readRawPart("x", `<x:docPr xmlns:x="urn:x" ${attributes}/>`);
      if (part === null) {
        throw new Error("fixture did not parse");
      }
      return objectFingerprint(part.root);
    };

    expect(fingerprint('name="Picture|title=Caption"')).not.toBe(
      fingerprint('name="Picture" title="Caption"'),
    );
    expect(fingerprint('name="Picture" title="Caption"')).toBe(
      fingerprint('title="Caption" name="Picture"'),
    );
  });
});

describe("emphasis for complex scripts", () => {
  test("bold for complex scripts travels on its own", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      p(
        `${r("complex only", "<w:bCs/>")}${r(" plain ")}${r("both", "<w:b/><w:bCs/>")}${r(" plain ")}${r("ordinary only", '<w:i/><w:iCs w:val="0"/>')}`,
      ),
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = (await readPart(docx, "word/document.xml")) ?? "";

    expect(document).toContain(
      '<w:r><w:rPr><w:bCs/></w:rPr><w:t xml:space="preserve">complex only</w:t></w:r>',
    );
    expect(document).toContain(
      '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">both</w:t></w:r>',
    );
    expect(document).toContain(
      '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">ordinary only</w:t></w:r>',
    );
    expect(report.emphasisSpans).toEqual({ source: 3, expected: 3, output: 3 });

    // Losing the complex-script half alone is a loss of emphasis.
    const plan = await planFor({ source, template });
    const lost = await verifyFor({
      source,
      template,
      plan,
      output: await replacePart(
        docx,
        "word/document.xml",
        document.replace("<w:rPr><w:bCs/></w:rPr>", ""),
      ),
    });
    expect(lost.issues.map(({ code }) => code)).toEqual(["emphasis-mismatch"]);
  });
});

describe("reference fields in place", () => {
  test("a field over other characters of the same text is caught", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      `${p(r("I and I"))}${p(r("Target"))}`,
    );
    const plan = await planFor({
      source,
      template,
      extra: {
        references: [
          {
            blockId: "p0",
            start: 0,
            end: 1,
            targetBlockId: "p1",
            instruction: "\\h",
            expectedText: "I",
          },
        ],
      },
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const field =
      /<w:r><w:fldChar w:fldCharType="begin"\/><\/w:r>.*?<w:r><w:fldChar w:fldCharType="end"\/><\/w:r>/u.exec(
        document,
      )?.[0] ?? "";
    const rest = '<w:r><w:t xml:space="preserve"> and I</w:t></w:r>';

    expect(report.references.results.map(({ status }) => status)).toEqual([
      "ok",
    ]);
    expect(document).toContain(`${field}${rest}`);

    const moved = document.replace(
      `${field}${rest}`,
      `<w:r><w:t xml:space="preserve">I and </w:t></w:r>${field}`,
    );
    const movedReport = await verifyFor({
      source,
      template,
      plan,
      output: await replacePart(docx, "word/document.xml", moved),
    });

    expect(unwrap(await analyzeDocxBlocks(docx)).blocks.at(0)?.text).toBe(
      "I and I",
    );
    expect(movedReport.issues.map(({ code }) => code)).toEqual([
      "reference-broken",
    ]);
    expect(movedReport.references.results.map(({ status }) => status)).toEqual([
      "missing-field",
    ]);
  });
});

describe("list instances from a base with its own overrides", () => {
  const OVERRIDE =
    '<w:lvlOverride w:ilvl="0" xmlns:x="urn:x"><w:lvl w:ilvl="0" x:tag="kept"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/><w:lvlJc w:val="left"/></w:lvl></w:lvlOverride>';

  const OTHER_LEVEL =
    '<w:lvlOverride w:ilvl="1" xmlns:y="urn:y"><w:lvl w:ilvl="1" y:tag="kept"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl></w:lvlOverride>';

  const build = async (restartFrom: number) => {
    const base = await buildTemplate();
    const numbering = (await readPart(base, "word/numbering.xml")) ?? "";
    const template = await replacePart(
      base,
      "word/numbering.xml",
      numbering.replace(
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/>${OVERRIDE}${OTHER_LEVEL}</w:num>`,
      ),
    );
    const source = await buildMinimalDocx(`${p(r("first"))}${p(r("second"))}`);
    const result = await applyFor({
      source,
      template,
      items: (blocks) =>
        blocks.map(({ id, hash }) => ({
          kind: "source",
          blockId: id,
          hash,
          styleId: "HouseBody",
          numbering: {
            numId: "new",
            baseNumId: 1,
            ilvl: 0,
            group: "list",
            restartFrom,
          },
        })),
    });
    const { docx, report } = unwrap(result);
    return {
      report,
      numbering: (await readPart(docx, "word/numbering.xml")) ?? "",
      labels: unwrap(await analyzeDocxBlocks(docx)).blocks.map(
        ({ label }) => label,
      ),
    };
  };

  test("a restart goes into the base override, which keeps its bindings", async () => {
    const { report, numbering, labels } = await build(3);

    expect(report.ok).toBe(true);
    expect(numbering).toContain(
      '<w:lvlOverride w:ilvl="0" xmlns:x="urn:x"><w:startOverride w:val="3"/><w:lvl w:ilvl="0" x:tag="kept">',
    );
    expect(readRawPart("word/numbering.xml", numbering)).not.toBeNull();
    expect(labels).toEqual(["3)", "4)"]);
  });

  test("an override of a level the list does not restart is copied whole", async () => {
    const { report, numbering } = await build(3);

    expect(report.ok).toBe(true);
    expect(numbering.split(OTHER_LEVEL)).toHaveLength(3);
  });
});

describe("freezing a simple field that binds a prefix", () => {
  const footer = (field: string, binding = ""): string =>
    `${XML}<w:ftr ${NS}${binding}>${p(`${r("File ")}${field}`)}</w:ftr>`;

  const freezeFooter = async (xml: string) => {
    const template = await replacePart(
      await buildTemplate(),
      "word/footer1.xml",
      xml,
    );
    const source = await buildMinimalDocx(p(r("Body text.")));
    const result = await applyFor({
      source,
      template,
      extra: { fields: { freezeKinds: ["FILENAME"] } },
    });
    return { result, template };
  };

  test("a prefix bound on the field keeps the field, and verification says so", async () => {
    const { result } = await freezeFooter(
      footer(
        '<w:fldSimple w:instr=" FILENAME " xmlns:x="urn:x"><w:r x:tag="1"><w:t>name.docx</w:t></w:r></w:fldSimple>',
      ),
    );

    expect(codeOf(result)).toBe("verification-failed");
    expect(issuesOf(result)).toEqual(["field-not-frozen"]);
  });

  test("a prefix bound on an ancestor leaves the result bound", async () => {
    const { result } = await freezeFooter(
      footer(
        '<w:fldSimple w:instr=" FILENAME "><w:r x:tag="1"><w:t>name.docx</w:t></w:r></w:fldSimple>',
        ' xmlns:x="urn:x"',
      ),
    );
    const { docx, report } = unwrap(result);
    const frozen = (await readPart(docx, "word/footer1.xml")) ?? "";

    expect(report.ok).toBe(true);
    expect(frozen).not.toContain("fldSimple");
    expect(frozen).toContain('<w:r x:tag="1"><w:t>name.docx</w:t></w:r>');
    expect(readRawPart("word/footer1.xml", frozen)).not.toBeNull();
  });
});

describe("a kept template block with a frozen field", () => {
  test("reads as the template has it", async () => {
    const template = await replacePart(
      await buildTemplate(),
      "word/document.xml",
      `${XML}<w:document ${NS}><w:body>${p(`${r("Dated ")}${fieldRuns("DATE", r("1.1.2020"))}`)}<w:sectPr/></w:body></w:document>`,
    );
    const source = await buildMinimalDocx(p(r("Body text.")));
    const { report } = unwrap(
      await applyFor({
        source,
        template,
        items: ([first]) => [
          { kind: "template", templateBlockIds: ["p0"] },
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
        extra: { fields: { freezeKinds: ["DATE"] } },
      }),
    );

    expect(report.ok).toBe(true);
    expect(report.issues).toEqual([]);
  });
});
