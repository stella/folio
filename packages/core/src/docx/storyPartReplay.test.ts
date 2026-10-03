import { panic } from "better-result";
import { expect, test } from "bun:test";
import JSZip from "jszip";

import { applyDocumentOps, paragraphLogicalText, type DocumentOp } from "@stll/docx-core/ops";
import type { Paragraph } from "../types/document";
import { visitDocxParagraphs } from "./paragraphTraversal";
import { parseDocx } from "./parser";
import { createDocx, repackDocx as repack } from "./rezip";

const repackDocx = (
  document: Parameters<typeof repack>[0],
  options: Parameters<typeof repack>[1] = {},
) => repack(document, { ...options, onDiagnostic: ({ type, part }) => panic(`${type}: ${part}`) });

const PROFILES = [
  {
    name: "transitional",
    word: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    relationships: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  },
  {
    name: "strict",
    word: "http://purl.oclc.org/ooxml/wordprocessingml/main",
    relationships: "http://purl.oclc.org/ooxml/officeDocument/relationships",
  },
] as const;

const CASES = PROFILES.flatMap((profile) =>
  (["header", "footer"] as const).flatMap((story) =>
    (["joinBlocks", "setParagraphProps"] as const).flatMap((operation) =>
      (["w", "alt"] as const).map((prefix) => ({ profile, story, operation, prefix })),
    ),
  ),
);

const NESTED_CASES = CASES.filter(({ operation }) => operation === "setParagraphProps").flatMap(
  (options) =>
    (["control", "table"] as const).map((target) => Object.assign({}, options, { target })),
);

type StoryFixtureOptions = Pick<(typeof CASES)[number], "profile" | "story" | "prefix">;
const storyFixture = async ({ profile, story, prefix: p }: StoryFixtureOptions) => {
  const root = story === "header" ? "hdr" : "ftr";
  const partPath = `word/${story}1.xml`;
  const paragraph = (id: string, text: string) =>
    `<${p}:p p14:paraId="${id}">\n  <${p}:r><${p}:rPr><${p}:sz ${p}:val="28"/><${p}:webHidden/></${p}:rPr><${p}:t>${text}</${p}:t></${p}:r>\n</${p}:p>`;
  const untouchedParagraph = paragraph("33333333", "Untouched");
  const controlSibling = paragraph("66666666", "Control sibling");
  const tableSibling = paragraph("77777777", "Cell sibling");
  const untouchedControl = `<${p}:sdt><${p}:sdtPr><${p}:id ${p}:val="9"/><${p}:richText/></${p}:sdtPr><${p}:sdtContent>\n${paragraph("44444444", "Nested control")}\n${controlSibling}\n</${p}:sdtContent></${p}:sdt>`;
  const untouchedTable = `<${p}:tbl><${p}:tblPr/><${p}:tblGrid><${p}:gridCol ${p}:w="2400"/></${p}:tblGrid><${p}:tr><${p}:tc><${p}:tcPr><${p}:tcW ${p}:w="2400" ${p}:type="dxa"/></${p}:tcPr>\n${paragraph("55555555", "Nested cell")}\n${tableSibling}\n</${p}:tc></${p}:tr></${p}:tbl>`;
  const storyXml = `<?xml version="1.0" encoding="UTF-8"?>\n<${p}:${root} xmlns:${p}="${profile.word}" xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml">\n<${p}:p p14:paraId="11111111"><${p}:pPr><${p}:keepNext/></${p}:pPr><${p}:r><${p}:rPr><${p}:sz ${p}:val="28"/><${p}:webHidden/></${p}:rPr><${p}:t>Target</${p}:t></${p}:r></${p}:p>\n<${p}:p p14:paraId="22222222"><${p}:r><${p}:t>Follower</${p}:t></${p}:r></${p}:p>\n${untouchedParagraph}\n${untouchedControl}\n${untouchedTable}\n</${p}:${root}>`;
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/${partPath}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${story}+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOffice" Type="${profile.relationships}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStory" Type="${profile.relationships}/${story}" Target="${story}1.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<${p}:document xmlns:${p}="${profile.word}" xmlns:rel="${profile.relationships}" xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml"><${p}:body><${p}:p p14:paraId="AAAAAAAA"><${p}:r><${p}:t>Main</${p}:t></${p}:r></${p}:p><${p}:sectPr><${p}:${story}Reference ${p}:type="default" rel:id="rIdStory"/></${p}:sectPr></${p}:body></${p}:document>`,
  );
  zip.file(partPath, storyXml);
  return {
    bytes: await zip.generateAsync({ type: "arraybuffer" }),
    partPath,
    storyXml,
    untouched: [untouchedParagraph, untouchedControl, untouchedTable],
    untouchedParagraph,
    nested: {
      control: { blockId: "44444444", sibling: controlSibling, otherContainer: untouchedTable },
      table: { blockId: "55555555", sibling: tableSibling, otherContainer: untouchedControl },
    },
  };
};

test.each(CASES)(
  "$story $operation preserves authored story bytes ($profile.name, $prefix prefix)",
  async ({ profile, story, operation, prefix }) => {
    const source = await storyFixture({ profile, story, prefix });
    const document = await parseDocx(source.bytes, { preloadFonts: false });
    const op = (
      operation === "joinBlocks"
        ? {
            type: operation,
            story: { kind: story, rId: "rIdStory" },
            blockId: "11111111",
            nextBlockId: "22222222",
            depth: 0,
          }
        : {
            type: operation,
            story: { kind: story, rId: "rIdStory" },
            blockId: "11111111",
            patch: { keepNext: false },
          }
    ) satisfies DocumentOp;
    const edited = applyDocumentOps(document, [op]).unwrap();
    const forward = await repackDocx(edited.document, { updateModifiedDate: false });
    const forwardZip = await JSZip.loadAsync(forward);
    const forwardXml = await forwardZip.file(source.partPath)?.async("text");
    for (const untouched of source.untouched) expect(forwardXml).toContain(untouched);

    const reopened = await parseDocx(forward, { preloadFonts: false });
    const parts = story === "header" ? reopened.package.headers : reopened.package.footers;
    const first = parts?.get("rIdStory")?.content.at(0);
    expect(first?.type).toBe("paragraph");
    if (first?.type !== "paragraph") throw new Error("Story fixture lost its edited paragraph");
    if (operation === "joinBlocks") expect(paragraphLogicalText(first)).toBe("TargetFollower");
    else expect(first.formatting?.keepNext).toBe(false);

    const restored = applyDocumentOps(edited.document, edited.inverse).unwrap().document;
    const inverseZip = await JSZip.loadAsync(
      await repackDocx(restored, { updateModifiedDate: false }),
    );
    expect(await inverseZip.file(source.partPath)?.async("uint8array")).toEqual(
      new TextEncoder().encode(source.storyXml),
    );
  },
);

test.each(NESTED_CASES)(
  "$story nested $target formatting preserves sibling bytes ($profile.name, $prefix prefix)",
  async ({ profile, story, prefix, target }) => {
    const source = await storyFixture({ profile, story, prefix });
    const nested = source.nested[target];
    const document = await parseDocx(source.bytes, { preloadFonts: false });
    const op = {
      type: "setParagraphProps",
      story: { kind: story, rId: "rIdStory" },
      blockId: nested.blockId,
      patch: { keepNext: false },
    } satisfies DocumentOp;
    const edited = applyDocumentOps(document, [op]).unwrap();
    const forward = await repackDocx(edited.document, { updateModifiedDate: false });
    const forwardZip = await JSZip.loadAsync(forward);
    const forwardXml = await forwardZip.file(source.partPath)?.async("text");
    for (const untouched of [source.untouchedParagraph, nested.otherContainer, nested.sibling])
      expect(forwardXml).toContain(untouched);

    const reopened = await parseDocx(forward, { preloadFonts: false });
    const matches: Paragraph[] = [];
    visitDocxParagraphs(
      {
        documentBody: reopened.package.document,
        ...(reopened.package.headers !== undefined ? { headers: reopened.package.headers } : {}),
        ...(reopened.package.footers !== undefined ? { footers: reopened.package.footers } : {}),
      },
      (paragraph) => {
        if (paragraph.paraId === nested.blockId) matches.push(paragraph);
      },
    );
    expect(matches).toHaveLength(1);
    expect(matches.at(0)?.formatting?.keepNext).toBe(false);

    const restored = applyDocumentOps(edited.document, edited.inverse).unwrap().document;
    const inverseZip = await JSZip.loadAsync(
      await repackDocx(restored, { updateModifiedDate: false }),
    );
    expect(await inverseZip.file(source.partPath)?.async("uint8array")).toEqual(
      new TextEncoder().encode(source.storyXml),
    );
  },
);

test("save paths forward source mismatch diagnostics with the resolved part path", async () => {
  const source = await storyFixture({ profile: PROFILES[0], story: "header", prefix: "w" });
  const document = await parseDocx(source.bytes, { preloadFonts: false });
  const header = document.package.headers?.get("rIdStory");
  if (!header) throw new Error("Missing header");
  header.verbatimFingerprint = "forged";
  const diagnostics: unknown[] = [];
  for (const save of [
    () =>
      repack(document, {
        updateModifiedDate: false,
        onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
      }),
    () => createDocx(document, { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) }),
  ]) {
    diagnostics.length = 0;
    const saved = await save();
    expect(diagnostics).toEqual([{ type: "sourceReplayMismatch", part: source.partPath }]);
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(reopened.package.headers?.get("rIdStory")?.content).toEqual(header.content);
  }
});
