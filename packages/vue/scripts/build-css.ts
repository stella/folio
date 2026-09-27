#!/usr/bin/env bun

// Keep @fontsource imports in the published stylesheet so consumer bundlers
// resolve the font files from this package's dependencies.
import { panic } from "better-result";
import { join } from "node:path";

const packageDir = join(import.meta.dir, "..");
const cssPath = join(packageDir, "dist", "folio-vue.css");
const stylesDir = join(packageDir, "src", "styles");
const fonts = await Bun.file(join(stylesDir, "fonts.css")).text();
const aliases = await Bun.file(join(stylesDir, "font-aliases.css")).text();
const editor = await Bun.file(cssPath).text();

if (!fonts.includes('@import "@fontsource/')) {
  panic("Vue stylesheet has no bundled font imports");
}

await Bun.write(cssPath, `${fonts}\n${aliases}\n${editor}`);
