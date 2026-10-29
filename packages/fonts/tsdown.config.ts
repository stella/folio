import { defineConfig } from "tsdown";

// `dist/index.js` and `dist/node.js` sit one level below the package root, as
// `src/` does, so the Node reader finds `../fonts/` from either.
export default defineConfig({
  entry: {
    index: "src/index.ts",
    node: "src/node.ts",
  },
  format: ["esm"],
  platform: "node",
  // `.js` like the other packages (the package is `"type": "module"`);
  // tsdown's Node default is `.mjs`, which prepare-publish would not find.
  fixedExtension: false,
  dts: { generator: "oxc" },
  outDir: "dist",
  clean: true,
});
