#!/usr/bin/env bun

// Keep the @stll/folio-fonts import in the published stylesheet so consumer
// bundlers resolve the font files from this package's dependencies.
import { panic } from "better-result";
import { join } from "node:path";

import { scopeEditorCss } from "../../../scripts/lib/scoped-editor-css";

const packageDir = join(import.meta.dir, "..");
const cssPath = join(packageDir, "dist", "folio-vue.css");
const stylesDir = join(packageDir, "src", "styles");
const fonts = await Bun.file(join(stylesDir, "fonts.css")).text();
const editor = scopeEditorCss(await Bun.file(cssPath).text());

if (!fonts.includes('@import "@stll/folio-fonts/fonts.css";')) {
  panic("Vue stylesheet has no bundled font imports");
}

await Bun.write(cssPath, `${fonts}\n${editor}`);
