/**
 * `w:sym` names both its attributes optionally, and folio refused the document.
 *
 * `CT_Sym` declares `w:font` and `w:char` as optional. Word opens a
 * `<w:sym w:char="F0B7"/>` by falling back to the run's font, and folio threw
 * on it twice over: the model validator called a missing font an error, and
 * the ProseMirror projection called a missing character one. A refusal is a
 * worse loss than a drop — the document does not open at all — so schema-valid
 * input has to reach the editor.
 *
 * The other half is not inventing what the document did not say: the parser
 * records an absent attribute as an empty string, and the serializer must
 * write that back as absence rather than as `w:font=""`.
 */

import { describe, expect, test } from "bun:test";

import { parseRun } from "../runParser";
import { serializeRun } from "../serializer/runSerializer";
import { parseXmlDocument, type XmlElement } from "../xmlParser";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const runFrom = (inner: string) => {
  const node = parseXmlDocument(`<w:r ${W_NS}>${inner}</w:r>`) as XmlElement | null;
  if (!node) {
    throw new Error("the fixture did not parse");
  }
  return parseRun(node, new Map(), null, null, null);
};

describe("w:sym keeps whichever of its optional attributes the document wrote", () => {
  test("a symbol with no font survives and gains none", () => {
    const run = runFrom('<w:sym w:char="F0B7"/>');
    const xml = serializeRun(run);
    expect(xml).toContain('<w:sym w:char="F0B7"/>');
    expect(xml).not.toContain("w:font");
  });

  test("a symbol with no character survives and gains none", () => {
    const run = runFrom('<w:sym w:font="Wingdings"/>');
    const xml = serializeRun(run);
    expect(xml).toContain('<w:sym w:font="Wingdings"/>');
    expect(xml).not.toContain("w:char");
  });

  test("a symbol with both is unchanged", () => {
    const run = runFrom('<w:sym w:font="Wingdings" w:char="F0B7"/>');
    expect(serializeRun(run)).toContain('<w:sym w:font="Wingdings" w:char="F0B7"/>');
  });
});
