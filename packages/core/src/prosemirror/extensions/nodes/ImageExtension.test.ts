import { describe, expect, test } from "bun:test";

import { schema } from "../../schema";

const isStringRecord = (value: unknown): value is Record<string, string> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

describe("ImageExtension border serialization", () => {
  test("serializes and parses borderStyle without using borderKind", () => {
    const image = schema.nodes.image.create({
      src: "data:image/png;base64,",
      width: 100,
      height: 80,
      borderWidth: 2,
      borderColor: "currentColor",
      borderStyle: "dashed",
    });

    const output = schema.nodes.image.spec.toDOM?.(image);
    expect(Array.isArray(output)).toBe(true);
    if (!Array.isArray(output)) {
      return;
    }
    expect(output.at(0)).toBe("img");

    const domAttrs = output.at(1);
    expect(isStringRecord(domAttrs)).toBe(true);
    if (!isStringRecord(domAttrs)) {
      return;
    }
    expect(domAttrs["data-border-width"]).toBe("2");
    expect(domAttrs["data-border-color"]).toBe("currentColor");
    expect(domAttrs["data-border-style"]).toBe("dashed");
    expect(domAttrs["data-border-kind"]).toBeUndefined();
    expect(domAttrs["style"]).toContain("border: 2px dashed currentColor");

    const parseRule = schema.nodes.image.spec.parseDOM?.at(0);
    expect(typeof parseRule?.getAttrs).toBe("function");
    if (typeof parseRule?.getAttrs !== "function") {
      return;
    }

    const parsedAttrs = parseRule.getAttrs({
      dataset: {
        borderWidth: domAttrs["data-border-width"],
        borderColor: domAttrs["data-border-color"],
        borderStyle: domAttrs["data-border-style"],
      },
      width: 100,
      height: 80,
      getAttribute(name: string): string | null {
        if (name === "src") {
          return "data:image/png;base64,";
        }
        return null;
      },
    } as unknown as HTMLElement);

    expect(parsedAttrs).toMatchObject({
      borderWidth: 2,
      borderColor: "currentColor",
      borderStyle: "dashed",
    });
  });
});
