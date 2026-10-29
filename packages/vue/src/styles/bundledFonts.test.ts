import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";

const readStylesheet = (packageName: "react" | "vue", file: string) =>
  readFileSync(
    new URL(`../../../${packageName}/src/styles/${file}`, import.meta.url),
    "utf8",
  );

const importSpecs = (css: string) =>
  [...css.matchAll(/@import "(?<spec>[^"]+)";/gu)].map(
    (match) => match.groups?.spec,
  );

test("Vue ships the same bundled font faces as React", () => {
  const vueFonts = readStylesheet("vue", "fonts.css");
  const reactFonts = readStylesheet("react", "fonts.css");
  const dependencies = JSON.parse(
    readFileSync(new URL("../../../vue/package.json", import.meta.url), "utf8"),
  ).dependencies;

  expect(importSpecs(vueFonts)).toEqual(["@stll/folio-fonts/fonts.css"]);
  expect(importSpecs(vueFonts)).toEqual(importSpecs(reactFonts));
  expect(dependencies).toHaveProperty("@stll/folio-fonts");
  expect(
    Object.keys(dependencies).filter((name) => name.startsWith("@fontsource/")),
  ).toEqual([]);
});
