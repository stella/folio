/**
 * The registry is the schema's mark set, or the tables checked against it are
 * checked against nothing.
 *
 * `SchemaMarkName` is `keyof typeof MARK_EXTENSIONS`, and a table that must
 * decide something per mark — what the painter does with it, what a
 * replacement does with it — is declared total over that union. The compiler
 * enforces the totality; these hold the union to the schema the editor
 * actually builds, so a mark added through another route cannot slip past the
 * tables by never joining the union.
 */

import { describe, expect, test } from "bun:test";

import { MARK_EXTENSIONS, MARK_NESTING_ORDER } from "./markRegistry";
import { schema } from "../schema";

describe("the mark registry", () => {
  test("names exactly the marks the schema declares", () => {
    expect(Object.keys(MARK_EXTENSIONS).toSorted()).toEqual(Object.keys(schema.marks).toSorted());
  });

  test("each entry builds the mark it is filed under", () => {
    for (const [name, extension] of Object.entries(MARK_EXTENSIONS)) {
      expect(extension().config.schemaMarkName).toBe(name);
    }
  });

  test("the nesting order ranks each mark once", () => {
    expect(new Set(MARK_NESTING_ORDER).size).toBe(MARK_NESTING_ORDER.length);
  });

  test("the nesting order is the schema's own mark order, which is DOM nesting order", () => {
    expect([...MARK_NESTING_ORDER]).toEqual(Object.keys(schema.marks));
  });

  test("the wrapper encloses an editor-created revision, which encloses the link", () => {
    const rank = (name: string): number => MARK_NESTING_ORDER.indexOf(name);

    expect(rank("inlineWrapper")).toBeLessThan(rank("insertion"));
    expect(rank("inlineWrapper")).toBeLessThan(rank("deletion"));
    expect(rank("insertion")).toBeLessThan(rank("hyperlink"));
    expect(rank("deletion")).toBeLessThan(rank("hyperlink"));
    expect(rank("inlineWrapper")).toBeLessThan(rank("hyperlink"));
  });
});
