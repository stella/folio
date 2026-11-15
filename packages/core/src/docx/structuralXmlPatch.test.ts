import { expect, test } from "bun:test";

import { wellFormedXmlIssue } from "@stll/docx-core/xml";

import { buildStructuralDocumentPatch } from "./structuralXmlPatch";
import { getChildElements, parseXmlDocument } from "./xmlParser";

type PatchOptions = Parameters<typeof buildStructuralDocumentPatch>[0];
/** The spliced part, or null where the splice declined. */
const splice = (options: PatchOptions): string | null => {
  const result = buildStructuralDocumentPatch(options);
  return result.type === "patched" ? result.xml : null;
};
/** Why the splice declined, or null where it spliced. */
const declined = (options: PatchOptions) => {
  const result = buildStructuralDocumentPatch(options);
  return result.type === "declined" ? result.reason : null;
};

const wrap = (body: string, prefix = "w") =>
  `<${prefix}:document xmlns:${prefix}="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:id="http://schemas.microsoft.com/office/word/2010/wordml"><${prefix}:body>${body}</${prefix}:body></${prefix}:document>`;
const paragraph = (id: number, text = `Paragraph ${id}`, prefix = "w") =>
  `<${prefix}:p id:paraId="${id.toString(16).padStart(8, "0")}"><${prefix}:r><${prefix}:t>${text}</${prefix}:t></${prefix}:r></${prefix}:p>`;
const ids = (xml: string) => {
  const root = parseXmlDocument(xml);
  const body = getChildElements(root).at(0);
  return getChildElements(body).map(
    (element) => element.attributes?.["id:paraId"],
  );
};

test("all deletion subsets and insertion gaps retain surviving source bytes and ordering", () => {
  const original = wrap(
    [1, 2, 3, 4].map((id) => paragraph(id)).join("\n<!-- untouched gap -->\n"),
  );
  for (let mask = 0; mask < 16; mask++) {
    const survivors = [1, 2, 3, 4].filter(
      (id) => (mask & (1 << (id - 1))) !== 0,
    );
    for (let gap = 0; gap <= survivors.length; gap++) {
      const desired = survivors.toSpliced(gap, 0, 5, 6, 7);
      const result = splice({
        originalXml: original,
        serializedXml: wrap(desired.map((id) => paragraph(id)).join("")),
        changedIds: new Set(),
      });
      expect(result).not.toBeNull();
      if (result === null) continue;
      expect(wellFormedXmlIssue(result)).toBeNull();
      expect(ids(result)).toEqual(
        desired.map((id) => id.toString(16).padStart(8, "0")),
      );
      for (const id of survivors) expect(result).toContain(paragraph(id));
      expect(result.match(/<!-- untouched gap -->/gu)?.length).toBe(3);
    }
  }
});

test("resolves aliased source elements and quoted delimiters without touching source gaps", () => {
  const original = wrap(
    `${paragraph(1, "one", "doc")}<!-- <doc:p> fake -->${paragraph(2, "two", "doc")}`,
    "doc",
  );
  const result = splice({
    originalXml: original,
    serializedXml: wrap(
      paragraph(1, "edited &gt; one") + paragraph(3) + paragraph(2),
    ),
    changedIds: new Set(["00000001"]),
  });
  expect(result).toContain(paragraph(2, "two", "doc"));
  expect(result).toContain("<!-- <doc:p> fake -->");
  expect(result).toContain(
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  );
  expect(result).toContain("edited &gt; one");
});

test("refuses moved, duplicate, missing, or reserved source identities", () => {
  for (const source of [paragraph(1) + paragraph(1), paragraph(0), "<w:p/>"]) {
    expect(
      splice({
        originalXml: wrap(source),
        serializedXml: wrap(paragraph(2)),
        changedIds: new Set(),
      }),
    ).toBeNull();
  }
  expect(
    splice({
      originalXml: wrap(paragraph(1) + paragraph(2)),
      serializedXml: wrap(paragraph(2) + paragraph(1)),
      changedIds: new Set(),
    }),
  ).toBeNull();
});

test("retains identical tables as barriers, writes an edited table whole, and refuses crossing them", () => {
  const table = `<w:tbl><w:tr><w:tc>${paragraph(9)}</w:tc></w:tr></w:tbl>`;
  const original = wrap(paragraph(1) + table + paragraph(2));
  const patch = (body: string) =>
    splice({
      originalXml: original,
      serializedXml: wrap(body),
      changedIds: new Set(),
    });
  expect(patch(paragraph(1) + paragraph(3) + table + paragraph(2))).toContain(
    table,
  );
  expect(patch(paragraph(2) + table + paragraph(1))).toBeNull();
  // An edited table is written from the serialized part, in its place.
  const edited = patch(
    paragraph(1) + table.replace("Paragraph 9", "edited") + paragraph(2),
  );
  expect(edited).toContain("edited");
  expect(edited).not.toContain("Paragraph 9");
  expect(edited?.indexOf("Paragraph 1")).toBeLessThan(
    edited?.indexOf("edited") ?? -1,
  );
  // A table added or removed still takes the full write.
  expect(patch(paragraph(1) + paragraph(2))).toBeNull();
  expect(patch(paragraph(1) + table + paragraph(9))).toBeNull();
});

test("refuses relationship payloads, section endpoints and Strict splices", () => {
  for (const element of ["sectPr", "drawing", "footnoteReference"]) {
    const original = wrap(
      paragraph(1).replace("</w:p>", `<w:${element}/></w:p>`),
    );
    expect(
      declined({
        originalXml: original,
        serializedXml: wrap(paragraph(1) + paragraph(2)),
        changedIds: new Set(["00000001"]),
      }),
    ).toBe("structuralDependentContent");
  }
  expect(
    declined({
      originalXml: wrap(paragraph(1)).replace(
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
      ),
      serializedXml: wrap(paragraph(1) + paragraph(2)),
      changedIds: new Set(),
    }),
  ).toBe("structuralUnreadableBody");
});

test("untouched image and section paragraphs remain verbatim beside inserted paragraphs", () => {
  for (const element of ["drawing", "sectPr"]) {
    const kept = paragraph(1).replace("</w:p>", `<w:${element}/></w:p>`);
    const result = splice({
      originalXml: wrap(kept),
      serializedXml: wrap(kept + paragraph(2)),
      changedIds: new Set(["00000002"]),
    });
    expect(result).toContain(kept);
    expect(result).toContain("Paragraph 2");
  }
});

test("splices a body insertion enclosed by a bookmark whose endpoints sit in separate tables", () => {
  const table = (id: number, marker: string) =>
    `<w:tbl><w:tr><w:tc>${paragraph(id).replace("</w:p>", marker + "</w:p>")}</w:tc></w:tr></w:tbl>`;
  const start = table(1, '<w:bookmarkStart w:id="1" w:name="b"/>');
  const end = table(2, '<w:bookmarkEnd w:id="1"/>');
  const result = splice({
    originalXml: wrap(start + end),
    serializedXml: wrap(start + paragraph(3) + end),
    changedIds: new Set(["00000003"]),
  });
  expect(result).toContain(start);
  expect(result).toContain(end);
  expect(result?.indexOf("Paragraph 3")).toBeGreaterThan(
    result?.indexOf("bookmarkStart") ?? Number.POSITIVE_INFINITY,
  );
  expect(result?.indexOf("Paragraph 3")).toBeLessThan(
    result?.indexOf("bookmarkEnd") ?? -1,
  );
});

const marked = (id: number, before: string, after = "") =>
  paragraph(id)
    .replace("<w:r>", `${before}<w:r>`)
    .replace("</w:p>", `${after}</w:p>`);
const FIELD_BEGIN =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const FIELD_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

test("splices with a bookmark elsewhere, keeping its paragraphs as source bytes", () => {
  // Bookmarks, a field, a permission range and a move range in paragraphs
  // the edit does not touch: none of them is split by an insertion elsewhere.
  const untouched = [
    marked(
      1,
      '<w:bookmarkStart w:id="0" w:name="_Toc1"/>',
      `${FIELD_BEGIN}<w:r><w:t>1</w:t></w:r>${FIELD_END}<w:bookmarkEnd w:id="0"/>`,
    ),
    marked(
      2,
      '<w:permStart w:id="4" w:edGrp="everyone"/>',
      '<w:permEnd w:id="4"/>',
    ),
    marked(
      4,
      '<w:moveFromRangeStart w:id="5" w:name="move1" w:author="A" w:date="2026-01-01T00:00:00Z"/>',
      '<w:moveFromRangeEnd w:id="5"/>',
    ),
  ];
  // The previous check declined on any of these anywhere in the body; each
  // is here, so restoring it fails every case below.
  const source = wrap(untouched.join("<!-- kept -->"));
  for (const kind of [
    "bookmarkStart",
    "bookmarkEnd",
    "permStart",
    "permEnd",
    "fldChar",
    "moveFromRangeStart",
    "moveFromRangeEnd",
  ])
    expect(source).toContain(`<w:${kind} `);
  for (let gap = 0; gap <= untouched.length; gap++) {
    const desired = untouched.toSpliced(gap, 0, paragraph(3));
    const result = splice({
      originalXml: source,
      serializedXml: wrap(desired.join("")),
      changedIds: new Set(["00000003"]),
    });
    expect(result).not.toBeNull();
    if (result === null) continue;
    expect(wellFormedXmlIssue(result)).toBeNull();
    for (const kept of untouched) expect(result).toContain(kept);
    expect(result).toContain("Paragraph 3");
  }
});

test("splices an insertion inside a range or field that crosses paragraphs", () => {
  for (const [open, close] of [
    ['<w:bookmarkStart w:id="1" w:name="b"/>', '<w:bookmarkEnd w:id="1"/>'],
    ['<w:permStart w:id="2" w:ed="someone"/>', '<w:permEnd w:id="2"/>'],
    [FIELD_BEGIN, FIELD_END],
  ] as const) {
    const first = marked(1, "", open);
    const last = marked(2, close);
    const inserted = splice({
      originalXml: wrap(first + "<!-- kept -->" + last),
      serializedXml: wrap(first + paragraph(3) + last),
      changedIds: new Set(["00000003"]),
    });
    expect(inserted).toContain(`${first}<!-- kept -->`);
    expect(inserted).toContain(last);
    expect(ids(inserted ?? "")).toEqual(["00000001", "00000003", "00000002"]);
    // A rewritten paragraph carries its end of the range as serialized.
    const edited = last.replace("Paragraph 2", "edited");
    const rewritten = splice({
      originalXml: wrap(first + "<!-- kept -->" + last),
      serializedXml: wrap(first + paragraph(3) + edited),
      changedIds: new Set(["00000002", "00000003"]),
    });
    expect(rewritten).toContain(`${first}<!-- kept -->`);
    expect(rewritten).toContain("edited");
    expect(rewritten).not.toContain("Paragraph 2");
    expect(ids(rewritten ?? "")).toEqual(["00000001", "00000003", "00000002"]);
  }
});

test("declines, naming the split, where an untouched paragraph would keep a moved range or field end", () => {
  for (const [open, close] of [
    ['<w:bookmarkStart w:id="1" w:name="b"/>', '<w:bookmarkEnd w:id="1"/>'],
    ['<w:permStart w:id="2" w:ed="someone"/>', '<w:permEnd w:id="2"/>'],
    [
      '<w:moveToRangeStart w:id="3" w:name="m" w:author="A" w:date="2026-01-01T00:00:00Z"/>',
      '<w:moveToRangeEnd w:id="3"/>',
    ],
    [FIELD_BEGIN, FIELD_END],
  ] as const) {
    const first = marked(1, "", open);
    // The serialized body moved the end into the new paragraph, but the
    // paragraph that held it is not among the changed ones: its source bytes
    // would keep a second end.
    expect(
      declined({
        originalXml: wrap(first + marked(2, close)),
        serializedXml: wrap(first + marked(3, close) + paragraph(2)),
        changedIds: new Set(["00000003"]),
      }),
    ).toBe("structuralRangeSplit");
    // A removed paragraph takes a range's only end with it while the start
    // stays behind in an untouched one.
    expect(
      declined({
        originalXml: wrap(first + marked(2, close)),
        serializedXml: wrap(paragraph(1) + paragraph(3)),
        changedIds: new Set(["00000003"]),
      }),
    ).toBe("structuralRangeSplit");
  }
});

test("keeps body-level range markers as barriers", () => {
  const start = '<w:bookmarkStart w:id="1" w:name="b"/>';
  const end = '<w:bookmarkEnd w:id="1"/>';
  const original = wrap(start + paragraph(1) + paragraph(2) + end);
  const result = splice({
    originalXml: original,
    serializedXml: wrap(
      start + paragraph(1) + paragraph(3) + paragraph(2) + end,
    ),
    changedIds: new Set(["00000003"]),
  });
  expect(result).toContain(start + paragraph(1));
  expect(result).toContain(paragraph(2) + end);
  expect(result).toContain("Paragraph 3");
  // A marker the serialized body no longer holds is not kept from the source.
  expect(
    declined({
      originalXml: original,
      serializedXml: wrap(paragraph(1) + paragraph(3) + paragraph(2)),
      changedIds: new Set(["00000003"]),
    }),
  ).toBe("structuralBarrierChanged");
});

test("splices beside comment anchors exactly when the result carries the serialized body's anchors", () => {
  const anchored = (id: number, marker: string) =>
    paragraph(id).replace("</w:p>", `${marker}</w:p>`);
  const start = anchored(1, '<w:commentRangeStart w:id="7"/>');
  const end = anchored(
    2,
    '<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>',
  );
  const patch = (serialized: string, changed: readonly string[] = []) =>
    splice({
      originalXml: wrap(`${start}<!-- kept -->${end}`),
      serializedXml: wrap(serialized),
      changedIds: new Set(changed),
    });
  // An insertion inside the commented range keeps both anchoring paragraphs
  // as their source bytes.
  const inserted = patch(start + paragraph(3) + end, ["00000003"]);
  expect(inserted).toContain(start);
  expect(inserted).toContain(end);
  expect(inserted).toContain("<!-- kept -->");
  expect(inserted).toContain("Paragraph 3");
  // So does one in a range that runs from one table into another.
  const table = (body: string) =>
    `<w:tbl><w:tr><w:tc>${body}</w:tc></w:tr></w:tbl>`;
  expect(
    splice({
      originalXml: wrap(table(start) + table(end)),
      serializedXml: wrap(table(start) + paragraph(3) + table(end)),
      changedIds: new Set(["00000003"]),
    }),
  ).toContain("Paragraph 3");
  // An anchor the serialized body moved out of an untouched paragraph would
  // stay behind in the source bytes: the splice is refused.
  expect(patch(paragraph(1) + paragraph(3) + end, ["00000003"])).toBeNull();
});

test("structural splices preserve valid locally bound replay paragraphs at every insertion gap", () => {
  const namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
  const mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const original = wrap(
    [1, 2, 3].map((id) => paragraph(id)).join("<!-- gap -->"),
  );
  for (const alias of ["mc", "compat"]) {
    for (let gap = 0; gap <= 3; gap++) {
      const desired = [1, 2, 3].toSpliced(gap, 0, 4);
      const localParagraph = (id: number) =>
        paragraph(id).replace(
          "<w:p ",
          `<w:p xmlns:w="${word}" xmlns:id="${namespace}" xmlns:${alias}="${mc}" ${alias}:Ignorable="id" `,
        );
      const serializedXml = wrap(desired.map(localParagraph).join(""));
      expect(wellFormedXmlIssue(serializedXml)).toBeNull();
      const result = splice({
        originalXml: original,
        serializedXml,
        changedIds: new Set(["00000002"]),
      });
      expect(result).not.toBeNull();
      if (result === null) continue;
      expect(wellFormedXmlIssue(result)).toBeNull();
      expect(ids(result)).toEqual(
        desired.map((id) => id.toString(16).padStart(8, "0")),
      );
      expect(result).toContain(paragraph(1));
      expect(result).toContain(paragraph(3));
      expect(result).toContain(`${alias}:Ignorable="id"`);
    }
  }
});

test("reads a source comment anchor's unprefixed id as the reader does", () => {
  const anchored = (id: number, marker: string) =>
    paragraph(id).replace("</w:p>", `${marker}</w:p>`);
  const start = anchored(1, '<w:commentRangeStart id="7"/>');
  const end = anchored(2, '<w:commentRangeEnd id="7"/>');
  const serializedStart = anchored(1, '<w:commentRangeStart w:id="7"/>');
  const serializedEnd = anchored(2, '<w:commentRangeEnd w:id="7"/>');
  const inserted = splice({
    originalXml: wrap(start + end),
    serializedXml: wrap(serializedStart + paragraph(3) + serializedEnd),
    changedIds: new Set(["00000003"]),
  });
  expect(inserted).toContain(start);
  expect(inserted).toContain(end);
  expect(inserted).toContain("Paragraph 3");
});
