/**
 * `w:cr` and `w:br` render the same line break but are distinct OOXML
 * elements. The editor must not canonicalize an authored carriage return into
 * a break merely because they share layout semantics.
 */

import { describe, expect, test } from "bun:test";

import type { Document, Run } from "../../types/document";
import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { parseRun } from "../runParser";
import { serializeRun } from "../serializer/runSerializer";
import { parseXmlDocument, type XmlElement } from "../xmlParser";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const parsedCarriageReturn = (): Run => {
  const element = parseXmlDocument(`<w:r ${W_NS}><w:cr/></w:r>`) as XmlElement | null;
  if (!element) {
    throw new Error("the carriage-return fixture did not parse");
  }
  return parseRun(element, null, null);
};

const firstRun = (document: Document): Run => {
  const paragraph = document.package.document.content.at(0);
  const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
  if (run?.type !== "run") {
    throw new Error("the carriage-return fixture did not produce a run");
  }
  return run;
};

describe("w:cr round-trip", () => {
  test("keeps its element identity through the editor projection", () => {
    const source: Document = {
      package: {
        document: {
          content: [{ type: "paragraph", content: [parsedCarriageReturn()] }],
        },
      },
    };

    const run = firstRun(fromProseDoc(toProseDoc(source), source));
    expect(run.content).toEqual([{ type: "break", sourceElement: "cr" }]);
    expect(serializeRun(run)).toContain("<w:cr/>");
    expect(serializeRun(run)).not.toContain("<w:br");
  });
});
