import { defineConfig } from "tsdown";

export default defineConfig({
  entry: {
    index: "src/runtime.ts",
    oxlint: "src/oxlint.ts",
    testing: "src/testing.ts",
  },
  format: ["esm"],
  platform: "neutral",
  dts: { generator: "oxc" },
  outDir: "dist",
  clean: true,
});
