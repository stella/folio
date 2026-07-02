import { describe, expect, test } from "bun:test";

import { findUncoveredUtilities } from "./standalone-css-coverage";

// A miniature packed standalone.css: scoped utilities (`.folio-root .<util>`),
// low-specificity token fallbacks, and a couple of hand-written doc classes.
const STANDALONE_CSS = `
@import "@fontsource/carlito/400.css";
:where(.folio-root){--popover:#fff;--muted:#eee}
.folio-root .flex{display:flex}
.folio-root .bg-popover{background-color:var(--popover)}
.folio-root .text-muted-foreground{color:var(--muted-foreground)}
.folio-root .hover\\:bg-muted\\/80:hover{background-color:var(--muted)}
.folio-root .h-\\[var\\(--x\\)\\]{height:var(--x)}
.folio-root .\\[\\&_svg\\]\\:text-destructive svg{color:red}
.folio-root .antialiased{-webkit-font-smoothing:antialiased}
.docx-insertion{color:green}
.folio-default-button{border:1px solid}
`;

const js = (code: string) => [{ file: "index.js", code }];

describe("findUncoveredUtilities", () => {
  test("passes when every className utility has a generated rule", () => {
    const result = findUncoveredUtilities({
      jsFiles: js(
        'jsx("div",{className:"flex bg-popover text-muted-foreground hover:bg-muted/80"})',
      ),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.uncovered).toEqual([]);
    expect(result.candidates).toBeGreaterThan(0);
  });

  test("resolves utilities from cn()/clsx() arguments", () => {
    const result = findUncoveredUtilities({
      jsFiles: js('className: cn("flex antialiased", active && "bg-popover", "h-[var(--x)]")'),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.uncovered).toEqual([]);
  });

  test("reports a utility used in JS but absent from the sheet (the drift it guards)", () => {
    const result = findUncoveredUtilities({
      jsFiles: js('className:"flex mt-[12345px]"'),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.uncovered).toEqual(["mt-[12345px]"]);
  });

  test("arbitrary-value and arbitrary-variant utilities match through CSS escaping", () => {
    const result = findUncoveredUtilities({
      jsFiles: js('className:"h-[var(--x)] [&_svg]:text-destructive"'),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.uncovered).toEqual([]);
  });

  test("extracts literals nested inside template-literal interpolations", () => {
    // The dist bundle keeps conditionals inside className templates; the
    // branches' classes must count, or a compile miss would ship silently.
    const result = findUncoveredUtilities({
      jsFiles: js('className: `flex antialiased ${on ? "bg-popover" : "text-muted-foreground"}`'),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.candidates).toBe(4);
    expect(result.uncovered).toEqual([]);
  });

  test("comparison operands inside interpolations are not class tokens", () => {
    // `position === "footer"` is a discriminator, not a class; only the branch
    // literal counts.
    const result = findUncoveredUtilities({
      jsFiles: js('className: `flex${position === "footer" ? " bg-popover" : ""}`'),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.candidates).toBe(2);
    expect(result.uncovered).toEqual([]);
  });

  test("reports drift for an uncovered utility inside an interpolation", () => {
    const result = findUncoveredUtilities({
      jsFiles: js('className: `flex ${on ? "mt-[12345px]" : "bg-popover"}`'),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.uncovered).toEqual(["mt-[12345px]"]);
  });

  test("an interpolation splitting a class never tokenizes as a complete class", () => {
    // `w-[${size}px]` is runtime-assembled; its static fragments must not be
    // reported (no static sheet could ever cover them).
    const result = findUncoveredUtilities({
      jsFiles: js("className: `flex w-[${size}px]`"),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.uncovered).toEqual([]);
  });

  test("ignores import specifiers, i18n keys, and element/attribute names", () => {
    // None of these appear in a className/cn context, so none are candidates.
    const result = findUncoveredUtilities({
      jsFiles: js(
        'import x from "@stll/folio-core/prosemirror/schema";\n' +
          't("acceptChange"); jsx("svg",{role:"img","aria-label":"x"});\n' +
          'const mode="all-markup";',
      ),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.candidates).toBe(0);
    expect(result.uncovered).toEqual([]);
  });

  test("skips folio's own semantic classes and inert/marker classes", () => {
    const result = findUncoveredUtilities({
      jsFiles: js(
        'className: cn("folio-root folio-editor docx-insertion", "group peer", "animate-in slide-in-from-right")',
      ),
      standaloneCss: STANDALONE_CSS,
    });
    expect(result.uncovered).toEqual([]);
  });
});
