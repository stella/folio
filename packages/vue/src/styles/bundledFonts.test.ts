import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { FONTSOURCE_PACKAGES } from "@stll/folio-core/fonts/fontsourceFaces";

const readStylesheet = (packageName: "react" | "vue", file: string) =>
  readFileSync(new URL(`../../../${packageName}/src/styles/${file}`, import.meta.url), "utf8");

const importSpecs = (css: string) =>
  [...css.matchAll(/@import "(?<spec>@fontsource\/[^"]+)";/gu)].map((match) => match.groups?.spec);

const faceRules = (css: string) =>
  [...css.matchAll(/@font-face\s*\{[^}]+\}/gu)].map((match) => match[0].replace(/\s+/gu, " "));

test("Vue ships the same bundled font faces and aliases as React", () => {
  const vueFonts = readStylesheet("vue", "fonts.css");
  const reactFonts = readStylesheet("react", "fonts.css");
  const imports = importSpecs(vueFonts);
  const dependencies = JSON.parse(
    readFileSync(new URL("../../../vue/package.json", import.meta.url), "utf8"),
  ).dependencies;

  expect(imports.length).toBeGreaterThan(0);
  expect(imports).toEqual(importSpecs(reactFonts));
  expect(new Set(imports.map((spec) => spec?.split("/").at(1)))).toEqual(
    new Set(Object.values(FONTSOURCE_PACKAGES)),
  );
  for (const packageName of Object.values(FONTSOURCE_PACKAGES)) {
    expect(dependencies).toHaveProperty(`@fontsource/${packageName}`);
  }
  const aliases = faceRules(readStylesheet("vue", "font-aliases.css"));
  expect(aliases.length).toBeGreaterThan(0);
  expect(aliases).toEqual(faceRules(readStylesheet("react", "font-aliases.css")));
});
