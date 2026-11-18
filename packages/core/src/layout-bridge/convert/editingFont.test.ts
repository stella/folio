import { describe, expect, test } from "bun:test";
import { DOMSerializer, Fragment } from "prosemirror-model";

import { acquireHarnessDom } from "../../../../../test/harnessDom";
import { schema } from "../../prosemirror/schema";
import { resolveFontFamily } from "../../utils/fontResolver";
import { editingFontFamily } from "./editingFont";

const fontMark = (attrs: Record<string, string>) =>
  schema.marks["fontFamily"]?.create(attrs) ??
  (() => {
    throw new TypeError("Missing fontFamily mark");
  })();

const families = (stack: string) =>
  stack
    .split(",")
    .map((family) => family.trim())
    .filter((family) => family !== "sans-serif");

describe("editingFontFamily", () => {
  test("draws a run with its Western stack, then its East Asian stack, as the painter resolves each slot", () => {
    const embedded = new Map([["MS Gothic", "folio-embedded-gothic"]]);
    const stack = editingFontFamily(
      fontMark({ ascii: "Calibri", hAnsi: "Calibri", eastAsia: "MS Gothic" }),
      { embeddedFontFamilyMap: embedded },
    );
    const western = families(
      resolveFontFamily("Calibri", { embeddedFontFamilyMap: embedded })
        .cssFallback,
    );
    const eastAsia = families(
      resolveFontFamily("MS Gothic", { embeddedFontFamilyMap: embedded })
        .cssFallback,
    ).filter((family) => !western.includes(family));
    expect(stack).toBe([...western, ...eastAsia, "sans-serif"].join(", "));
    expect(eastAsia.at(0)).toContain("folio-embedded-gothic");
  });

  test("resolves theme slots through the document's theme", () => {
    const stack = editingFontFamily(
      fontMark({ ascii: "Calibri", asciiTheme: "majorHAnsi" }),
      { theme: { fontScheme: { majorFont: { latin: "Aptos" } } } },
    );
    expect(stack?.startsWith(`"Aptos"`) || stack?.startsWith("Aptos")).toBe(
      true,
    );
  });

  test("a mark that names no font draws with the surrounding font", () => {
    expect(editingFontFamily(fontMark({}), {})).toBeNull();
  });

  test("the mark's DOM rendering takes the same stack", () => {
    const releaseDom = acquireHarnessDom();
    try {
      const mark = fontMark({ ascii: "Calibri", eastAsia: "MS Gothic" });
      const container = document.createElement("div");
      container.append(
        DOMSerializer.fromSchema(schema).serializeFragment(
          Fragment.from(schema.text("あ", [mark])),
          { document },
        ),
      );
      const probe = document.createElement("span");
      probe.style.fontFamily = editingFontFamily(mark, {}) ?? "";
      const span = container.querySelector("span");
      expect(span?.style.fontFamily).toBe(probe.style.fontFamily);
      expect(span?.style.fontFamily).toContain("MS Gothic");
    } finally {
      releaseDom();
    }
  });
});
