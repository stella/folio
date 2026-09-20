/**
 * The modelled interactive state reaches the control-kind element.
 *
 * `w:date`, `w:dropDownList`, `w:comboBox` and `w14:checkbox` are kept as
 * bytes so their unmodelled children and attributes survive, but the four
 * values a user changes are modelled. These tests pin the merge: the model
 * wins for those four, and everything else in the element is untouched —
 * including a producer's own namespace prefix, which the rewrite must not
 * swap for a foreign one.
 */

import { describe, expect, test } from "bun:test";

import { withModelledControlState } from "./sdtPropertiesPatch";

const checkbox = (checked: boolean) => ({ sdtType: "checkbox", checked }) as const;

describe("withModelledControlState — checkbox state", () => {
  test("updates an existing w14:checked attribute when the user toggles on", () => {
    const out = withModelledControlState(
      '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612"/></w14:checkbox>',
      checkbox(true),
    );
    expect(out).toContain('<w14:checked w14:val="1"/>');
    expect(out).not.toContain('w14:val="0"');
    // The glyph the author chose is not the state, and survives.
    expect(out).toContain('<w14:checkedState w14:val="2612"/>');
  });

  test("updates an existing w:checked attribute (no w14: prefix variant)", () => {
    expect(
      withModelledControlState('<w:checkbox><w:checked w:val="0"/></w:checkbox>', checkbox(true)),
    ).toContain('<w:checked w:val="1"/>');
  });

  test("injects a w14:checked when only the wrapper exists", () => {
    expect(withModelledControlState("<w14:checkbox></w14:checkbox>", checkbox(true))).toContain(
      '<w14:checked w14:val="1"/>',
    );
  });

  test("replaces the expanded `<w14:checked></w14:checked>` form without leaving stray closing tags", () => {
    // OOXML lets the empty element be written self-closing or expanded.
    // The expanded one is pinned here so the rewrite does not produce
    // `<w14:checked .../></w14:checked>`.
    const out = withModelledControlState(
      '<w14:checkbox><w14:checked w14:val="0"></w14:checked></w14:checkbox>',
      checkbox(true),
    );
    expect(out).toContain('<w14:checked w14:val="1"/>');
    expect(out).not.toMatch(/<w14:checked[^>]*\/>\s*<\/w14:checked>/u);
    expect(out).not.toContain('w14:val="0"');
  });
});

describe("withModelledControlState — date", () => {
  test("replaces an expanded-empty dateFormat element on round-trip", () => {
    // A producer that writes `<w:dateFormat …></w:dateFormat>` instead of
    // self-closing would otherwise leave a stale sibling beside the fresh
    // replacement: Word would see two dateFormat children and the picked
    // display string would not stick.
    const out = withModelledControlState(
      '<w:date w:fullDate="2026-06-02"><w:dateFormat w:val="d MMMM yyyy"></w:dateFormat></w:date>',
      { sdtType: "date", dateFormat: "yyyy-MM-dd", dateValueISO: "2026-06-02" },
    );
    expect(out.match(/dateFormat/giu)?.length).toBe(1);
    expect(out).toContain('<w:dateFormat w:val="yyyy-MM-dd"/>');
    expect(out).not.toContain('w:val="d MMMM yyyy"');
  });

  test("updates the format without disturbing the rest of the w:date element", () => {
    const out = withModelledControlState(
      '<w:date w:fullDate="2026-06-02"><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="en-GB"/><w:calendar w:val="gregorian"/></w:date>',
      { sdtType: "date", dateFormat: "yyyy-MM-dd", dateValueISO: "2026-06-02" },
    );
    expect(out).toContain('<w:dateFormat w:val="yyyy-MM-dd"/>');
    // The locale and the calendar are the element's own; nothing models them.
    expect(out).toContain('<w:lid w:val="en-GB"/>');
    expect(out).toContain('<w:calendar w:val="gregorian"/>');
  });
});

describe("withModelledControlState — dropdown last value", () => {
  test("strips an existing lastValue under a non-`w` source prefix and re-emits with the same prefix", () => {
    // A literal " w:lastValue=" marker misses `ns0:lastValue` entirely, so
    // the rewrite would append a fresh one beside the stale one and Word
    // would be free to keep reading the old value.
    const out = withModelledControlState(
      '<ns0:dropDownList ns0:lastValue="old"><ns0:listItem ns0:displayText="A" ns0:value="a"/><ns0:listItem ns0:displayText="B" ns0:value="b"/></ns0:dropDownList>',
      { sdtType: "dropdown", dropdownLastValue: "b" },
    );
    expect(out).not.toContain('lastValue="old"');
    expect(out).toContain('ns0:lastValue="b"');
    expect(out.match(/lastValue=/giu)?.length).toBe(1);
    expect(out).toContain('ns0:value="a"');
  });

  test("rewrites w:lastValue when the user picks a new item", () => {
    const out = withModelledControlState(
      '<w:dropDownList w:lastValue="ca"><w:listItem w:displayText="California" w:value="ca"/><w:listItem w:displayText="New York" w:value="ny"/></w:dropDownList>',
      { sdtType: "dropdown", dropdownLastValue: "ny" },
    );
    expect(out).toContain('w:lastValue="ny"');
    expect(out).not.toContain('w:lastValue="ca"');
    expect(out).toContain('w:value="ca"');
    expect(out).toContain('w:value="ny"');
  });

  test("escapes special characters in the new last value", () => {
    expect(
      withModelledControlState("<w:dropDownList/>", {
        sdtType: "dropdown",
        dropdownLastValue: 'Q&A "wrapped"',
      }),
    ).toContain('w:lastValue="Q&amp;A &quot;wrapped&quot;"');
  });
});

describe("withModelledControlState — a kind that states nothing interactive", () => {
  test("leaves the element alone", () => {
    const text = '<w:text w:multiLine="1"/>';
    expect(withModelledControlState(text, { sdtType: "plainText" })).toBe(text);
    const gallery = '<w:docPartObj><w:docPartGallery w:val="Quick Parts"/></w:docPartObj>';
    expect(withModelledControlState(gallery, { sdtType: "buildingBlockGallery" })).toBe(gallery);
  });
});
