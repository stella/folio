import { describe, expect, test } from "bun:test";

import { compareSide, withoutCarriers } from "./corpus-baseline-respell";

const EDITOR_ROUND_TRIP = "editor-round-trip";

const row = (message: string, files = 1) => ({ invariant: EDITOR_ROUND_TRIP, message, files });

const PREFIX = "editor round trip changed";

describe("the carriers come back out of a message", () => {
  test("only inside the model path", () => {
    expect(
      withoutCarriers(
        `${PREFIX} package.document.content[run].preservedAttributes: array [1] absent`,
      ),
    ).toBe(`${PREFIX} package.document.content[].preservedAttributes: array [1] absent`);
  });

  test("a message with no model path is left alone", () => {
    expect(withoutCarriers("DocxParseError: part word/document.xml [truncated]")).toBe(
      "DocxParseError: part word/document.xml [truncated]",
    );
  });
});

/**
 * Every row moves when a segment starts naming its carrier, so the diff of a
 * re-measured baseline is thousands of lines in which a new defect and a fixed
 * one look exactly like the churn. What the comparison has to answer is
 * narrower: is anything here besides the spelling?
 */
describe("what a re-spelling explains", () => {
  test("one recorded row split across two carriers is explained", () => {
    const recorded = [
      row(`${PREFIX} package.document.content[].preservedAttributes: array became absent`, 9),
    ];
    const observed = [
      row(`${PREFIX} package.document.content[run].preservedAttributes: array became absent`, 7),
      row(
        `${PREFIX} package.document.content[paragraph].preservedAttributes: array became absent`,
        4,
      ),
    ];
    expect(compareSide(recorded, observed)).toEqual({
      exact: 2,
      byPrefix: 0,
      appeared: [],
      vanished: [],
    });
  });

  test("an observed row nothing recorded explains is reported, not counted", () => {
    const recorded = [row(`${PREFIX} package.document.content[].bold: true became false`)];
    const observed = [row(`${PREFIX} package.document.content[run].italic: true became false`)];
    const { appeared, vanished } = compareSide(recorded, observed);
    expect(appeared).toEqual(observed);
    expect(vanished).toEqual(recorded);
  });

  /**
   * The old cap cut a message from the right, so a recorded row that lost its
   * tail is a prefix of the row it grew into, and the two are one defect.
   */
  test("a recorded row the old cap cut matches the row it grew into", () => {
    const recorded = [row(`${PREFIX} package.document.content[].rows[].cells[].conte…`)];
    const observed = [
      row(
        `${PREFIX} package.document.content[table].rows[tableRow].cells[tableCell].content[run].preservedAttributes: array became absent`,
      ),
    ];
    expect(compareSide(recorded, observed)).toEqual({
      exact: 0,
      byPrefix: 1,
      appeared: [],
      vanished: [],
    });
  });

  /**
   * A row that lost its middle is a prefix of nothing, so it is reported
   * rather than paired with whatever happens to share its head. A recorded row
   * predates the shortening, so this only fires when a path is deep enough
   * that even the new cap cannot hold it.
   */
  test("a row that lost its middle is reported, not paired by its head", () => {
    const recorded = [
      row(
        `${PREFIX} package.document.content[].rows[].cells[].content[].preservedAttributes: array became abse…`,
      ),
    ];
    const observed = [
      row(`${PREFIX} package.….content[run].preservedAttributes: array became absent`),
    ];
    expect(compareSide(recorded, observed)).toEqual({
      exact: 0,
      byPrefix: 0,
      appeared: observed,
      vanished: recorded,
    });
  });
});
