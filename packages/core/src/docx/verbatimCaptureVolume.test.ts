/**
 * Verbatim capture is proportional to the document, not to its nesting.
 *
 * A parser that keeps source markup beside the values it parsed can pay for
 * the same bytes once per level it captures at: capture a table, then each
 * row, then each cell, and a document nested N deep costs N times its own
 * size. folio does not do this — capture happens at the property elements
 * (`w:tblPr`, `w:trPr`, `w:tcPr`, `w:pPr`), which are disjoint siblings, never
 * at the containers that nest. Nothing in the code says so, though, and the
 * next capture site added at a container level would reintroduce it silently
 * and show up only as memory. So the ratio is measured here and bounded.
 *
 * The bound is on what the model retains, not on what capture produced:
 * every XML-looking string reachable from the parsed document is summed, so a
 * capture field added later is counted without this test being updated.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;

/** `depth` nested tables, two rows a table, paragraphs at the innermost cell. */
const nestedTables = (depth: number): string => {
  let inner = Array.from(
    { length: 2 },
    (_, index) =>
      `<w:p><w:pPr><w:pStyle w:val="s${String(index)}"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>cell ${String(index)}</w:t></w:r></w:p>`,
  ).join("");
  for (let level = 0; level < depth; level += 1) {
    const cell = `<w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${inner}</w:tc>`;
    const row = `<w:tr><w:trPr><w:trHeight w:val="20"/></w:trPr>${cell}</w:tr>`;
    inner = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>${row.repeat(2)}</w:tbl>`;
  }
  return `<?xml version="1.0"?><w:document ${NS}><w:body>${inner}<w:sectPr/></w:body></w:document>`;
};

const packageOf = async (documentXml: string): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types />");
  zip.file("word/document.xml", documentXml);
  return await zip.generateAsync({ type: "arraybuffer" });
};

/**
 * Every XML-looking string the parsed model retains.
 *
 * Deliberately structural rather than a list of capture fields: a list would
 * have to be kept in step with the parsers by hand, and a field it missed
 * would be exactly the field that regressed.
 */
const retainedMarkupChars = (root: unknown): number => {
  let total = 0;
  const seen = new WeakSet<object>();
  const pending: unknown[] = [root];
  while (pending.length > 0) {
    const current = pending.pop();
    if (typeof current === "string") {
      if (current.startsWith("<")) {
        total += current.length;
      }
      continue;
    }
    if (typeof current !== "object" || current === null || seen.has(current)) {
      continue;
    }
    seen.add(current);
    if (Array.isArray(current)) {
      for (const item of current) {
        pending.push(item);
      }
      continue;
    }
    if (current instanceof Map) {
      for (const value of current.values()) {
        pending.push(value);
      }
      continue;
    }
    for (const value of Object.values(current)) {
      pending.push(value);
    }
  }
  return total;
};

/**
 * Capture is allowed to cost a constant multiple of the document. The measured
 * value is around 0.5; the bound is slack enough that ordinary parser changes
 * do not trip it and tight enough that per-level re-capture, whose ratio grows
 * without limit, does.
 */
const MAX_CAPTURE_RATIO = 1.5;

describe("retained verbatim capture", () => {
  test("stays a constant multiple of the document as nesting deepens", async () => {
    const ratios: { depth: number; ratio: number }[] = [];
    for (const depth of [2, 4, 8, 12]) {
      const documentXml = nestedTables(depth);
      const document = await parseDocx(await packageOf(documentXml));
      const ratio = retainedMarkupChars(document) / documentXml.length;
      ratios.push({ depth, ratio });
      expect(ratio).toBeLessThan(MAX_CAPTURE_RATIO);
    }

    // The ratio must not trend with depth. Per-level re-capture would make the
    // deepest document's ratio a multiple of the shallowest one's; a constant
    // factor leaves them within measurement noise of each other.
    const shallow = ratios[0]?.ratio ?? 0;
    const deep = ratios.at(-1)?.ratio ?? 0;
    expect(shallow).toBeGreaterThan(0);
    expect(deep).toBeLessThan(shallow * 1.25);
  });
});
