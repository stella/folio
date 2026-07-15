import { describe, expect, test } from "bun:test";

import { ContentControlLockedError } from "../content-controls";
import {
  ContentControlWidgetController,
  parseContentControlListItems,
} from "./contentControlWidgetController";

const anchor = (left: number, bottom: number) => ({
  getBoundingClientRect: () => ({ bottom, left }),
});

describe("parseContentControlListItems", () => {
  test("keeps only valid display/value pairs", () => {
    expect(
      parseContentControlListItems(
        JSON.stringify([
          { displayText: "California", value: "ca" },
          { displayText: "Missing value" },
          null,
          { displayText: "New York", value: "ny" },
        ]),
      ),
    ).toEqual([
      { displayText: "California", value: "ca" },
      { displayText: "New York", value: "ny" },
    ]);
  });

  test("treats malformed input as an empty list", () => {
    expect(parseContentControlListItems("{")).toEqual([]);
    expect(parseContentControlListItems(undefined)).toEqual([]);
  });
});

describe("ContentControlWidgetController", () => {
  test("projects dropdown events into immutable picker state", () => {
    const controller = new ContentControlWidgetController();
    controller.handleWidgetEvent({
      kind: "dropdownOpen",
      tag: "state",
      pmPos: 17,
      sdtType: "dropdown",
      anchor: anchor(24, 80),
      listItemsJson: JSON.stringify([{ displayText: "New York", value: "ny" }]),
    });

    expect(controller.getSnapshot()).toEqual({
      status: "dropdown",
      items: [{ displayText: "New York", value: "ny" }],
      pmPos: 17,
      position: { x: 24, y: 84 },
    });
    controller.close();
    expect(controller.getSnapshot()).toEqual({ status: "closed" });
  });

  test("preserves the current date and surfaces typed refusals", () => {
    let refused = false;
    const controller = new ContentControlWidgetController({
      onRefused: () => {
        refused = true;
      },
    });
    controller.handleWidgetEvent({
      kind: "datePick",
      tag: "effective",
      pmPos: 9,
      anchor: anchor(10, 20),
      currentValue: "2026-06-02",
    });
    expect(controller.getSnapshot()).toEqual({
      status: "date",
      currentValue: "2026-06-02",
      pmPos: 9,
      position: { x: 10, y: 24 },
    });

    controller.handleWidgetEvent({
      kind: "refused",
      tag: "effective",
      pmPos: 9,
      sdtType: "date",
      anchor: anchor(10, 20),
      error: new ContentControlLockedError({ message: "locked", lock: "contentLocked" }),
    });
    expect(refused).toBe(true);
    expect(controller.getSnapshot()).toEqual({ status: "closed" });
  });
});
