import { describe, expect, test } from "bun:test";
import { DOMParser, DOMSerializer } from "prosemirror-model";

import { acquireHarnessDom } from "../../../../../../test/harnessDom";
import { expectImageAttrs } from "../../attrs";
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

// Earlier generated clipboard cases always supplied a drawing name, so they
// never exercised the schema's null default or distinguished an empty name.
test.each([undefined, null, "", "null", "Drawing name"])(
  "HTML clipboard keeps drawing name %j without inventing optional image metadata",
  (docPrName) => {
    const releaseDom = acquireHarnessDom();
    try {
      const image = schema.node("image", {
        src: "data:image/png;base64,",
        ...(docPrName === undefined ? {} : { docPrName }),
      });
      const host = document.createElement("div");
      host.append(
        DOMSerializer.fromSchema(schema).serializeNode(
          schema.node("paragraph", null, [image]),
        ),
      );
      const element = host.querySelector("img");
      if (element === null)
        throw new TypeError("Clipboard image has no HTML element.");
      expect(element.getAttribute("data-doc-pr-name")).toBe(docPrName ?? null);
      for (const attribute of [
        "alt",
        "title",
        "width",
        "height",
        "data-rid",
        "data-css-float",
        "data-transform",
        "data-opacity",
        "data-brightness",
        "data-contrast",
        "data-border-width",
        "data-border-color",
        "data-border-style",
        "data-picture-name",
        "data-picture-alt",
        "data-picture-title",
      ]) {
        expect(element.getAttribute(attribute)).toBeNull();
      }
      const parsed =
        DOMParser.fromSchema(schema).parseSlice(host).content.firstChild
          ?.firstChild;
      if (parsed?.type.name !== "image")
        throw new TypeError("Clipboard parse lost its image.");
      expect(expectImageAttrs(parsed).docPrName).toBe(docPrName ?? undefined);
    } finally {
      releaseDom();
    }
  },
);

describe("ImageExtension sources", () => {
  test("the editor DOM only loads local image sources, and the clipboard keeps the rest", () => {
    const releaseDom = acquireHarnessDom();
    try {
      const sources = [
        "data:image/png;base64,iVBORw0KGgo=",
        "blob:https://folio.test/7f3c",
        "https://images.example.test/pixel.png?t=1",
        "http://images.example.test/pixel.png",
        "//images.example.test/pixel.png",
      ];
      for (const src of sources) {
        const host = document.createElement("div");
        host.append(
          DOMSerializer.fromSchema(schema).serializeNode(
            schema.node("paragraph", null, [schema.node("image", { src })]),
          ),
        );
        const element = host.querySelector("img");
        if (element === null) throw new TypeError("Image has no element.");
        const local = src.startsWith("data:") || src.startsWith("blob:");
        expect(element.getAttribute("src")).toBe(local ? src : null);

        const parsed =
          DOMParser.fromSchema(schema).parseSlice(host).content.firstChild
            ?.firstChild;
        if (parsed?.type.name !== "image")
          throw new TypeError("Clipboard parse lost its image.");
        expect(expectImageAttrs(parsed).src).toBe(src);
      }
    } finally {
      releaseDom();
    }
  });
});
