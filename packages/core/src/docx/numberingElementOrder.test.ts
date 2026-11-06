/**
 * `w:numbering` holds its abstract definitions, then its instances, then
 * `w:numIdMacAtCleanup`. A definition added on save keeps that order.
 */

import { describe, expect, test } from "bun:test";

import { appendNumberingDefs } from "./selectiveXmlPatch";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const ABSTRACT = (id: number) =>
  `<w:abstractNum w:abstractNumId="${String(id)}"><w:multiLevelType w:val="hybridMultilevel"/></w:abstractNum>`;
const NUM = (id: number, abstractId: number) =>
  `<w:num w:numId="${String(id)}"><w:abstractNumId w:val="${String(abstractId)}"/></w:num>`;
const CLEANUP = '<w:numIdMacAtCleanup w:val="1"/>';
const part = (...children: string[]) =>
  `<w:numbering ${NS}>${children.join("")}</w:numbering>`;

const childOrder = (xml: string) =>
  [...xml.matchAll(/<w:(abstractNum|num|numIdMacAtCleanup)[ />]/gu)].map(
    ([, name]) => name,
  );

describe("a numbering definition added on save", () => {
  const current = part(ABSTRACT(1), ABSTRACT(2), NUM(1, 1), NUM(2, 2));
  const added = {
    abstractNums: new Set(["2"]),
    nums: new Set(["2"]),
  };

  test("stands before a part's cleanup marker", () => {
    const result = appendNumberingDefs(
      part(ABSTRACT(1), NUM(1, 1), CLEANUP),
      current,
      added,
    );
    expect(childOrder(result ?? "")).toEqual([
      "abstractNum",
      "abstractNum",
      "num",
      "num",
      "numIdMacAtCleanup",
    ]);
  });

  test("stands before the cleanup marker when the part has no instance", () => {
    const result = appendNumberingDefs(
      part(ABSTRACT(1), CLEANUP),
      current,
      added,
    );
    expect(childOrder(result ?? "")).toEqual([
      "abstractNum",
      "abstractNum",
      "num",
      "numIdMacAtCleanup",
    ]);
  });

  test("ends the part when there is no cleanup marker", () => {
    const result = appendNumberingDefs(
      part(ABSTRACT(1), NUM(1, 1)),
      current,
      added,
    );
    expect(childOrder(result ?? "")).toEqual([
      "abstractNum",
      "abstractNum",
      "num",
      "num",
    ]);
  });
});
