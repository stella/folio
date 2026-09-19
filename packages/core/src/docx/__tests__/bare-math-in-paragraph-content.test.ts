/**
 * A bare OMML element is paragraph content, and a tracked wrapper full of it
 * was written back empty.
 *
 * Every group that admits `m:oMath` also admits `m:EG_OMathMathElements`, so
 * `<w:ins><m:f/></w:ins>` is a tracked insertion of a fraction with no
 * `m:oMath` around it. The parser recognised only the two wrapper elements and
 * let the rest fall off the end of its `switch`, so the revision reached disk
 * with its content gone — a reviewer accepting an edit that is no longer
 * there, which is why this is an integrity bug rather than a fidelity one.
 *
 * The equation carries no structure the editable model holds, so it travels as
 * the markup it arrived as, exactly like the equations that do have a wrapper.
 */

import { describe, expect, test } from "bun:test";

import { parseParagraph } from "../paragraphParser";
import { serializeParagraph } from "../serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "../xmlParser";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

const paragraphFrom = (inner: string) => {
  const node = parseXmlDocument(`<w:p ${NS}>${inner}</w:p>`) as XmlElement | null;
  if (!node) {
    throw new Error("the fixture did not parse");
  }
  return parseParagraph(node, new Map(), null, null);
};

const TRACKED = 'w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"';

describe("a bare OMML element survives as paragraph content", () => {
  test("a tracked insertion keeps the fraction it inserted", () => {
    const paragraph = paragraphFrom(
      `<w:r><w:t>before</w:t></w:r><w:ins ${TRACKED}><m:f><m:num/><m:den/></m:f></w:ins>`,
    );
    const xml = serializeParagraph(paragraph);
    expect(xml).toContain("<m:f><m:num/><m:den/></m:f>");
    expect(xml).not.toContain(
      '<w:ins w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"></w:ins>',
    );
  });

  test("a bare accent outside any wrapper survives too", () => {
    const paragraph = paragraphFrom("<w:r><w:t>before</w:t></w:r><m:acc><m:e/></m:acc>");
    expect(serializeParagraph(paragraph)).toContain("<m:acc><m:e/></m:acc>");
  });

  test("an m:oMath wrapper is unchanged", () => {
    const paragraph = paragraphFrom("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>");
    expect(serializeParagraph(paragraph)).toContain("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>");
  });
});
