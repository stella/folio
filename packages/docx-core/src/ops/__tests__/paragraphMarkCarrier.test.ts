import { describe, expect, test } from "bun:test";

import { paragraphMarkCarrier } from "../blockDeletion";

describe("paragraphMarkCarrier", () => {
  test("the following paragraph owns a deleted break", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: "paragraph",
        endsContainer: false,
        removesCurrentMark: true,
        removesParagraph: false,
      }),
    ).toBe("following");
  });

  test("a container-final paragraph keeps its required terminator", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: undefined,
        endsContainer: true,
        removesCurrentMark: true,
        removesParagraph: false,
      }),
    ).toBe("current");
  });

  test("a cleared split paragraph's retained mark stays before a table", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: "table",
        endsContainer: false,
        removesCurrentMark: false,
        removesParagraph: false,
      }),
    ).toBe("current");
  });

  test("a paragraph mark remains the carrier before a table while its review resolves", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: "table",
        endsContainer: false,
        removesCurrentMark: true,
        removesParagraph: false,
      }),
    ).toBe("current");
  });

  test("a resolved inserted paragraph stays as the table's preceding carrier", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: "table",
        endsContainer: false,
        removesCurrentMark: true,
        removesParagraph: false,
      }),
    ).toBe("current");
  });

  test("a removed inserted paragraph needs no container-final carrier", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: undefined,
        endsContainer: true,
        removesCurrentMark: true,
        removesParagraph: true,
      }),
    ).toBe("none");
  });
});
