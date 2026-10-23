import { defineConfig } from "tsdown";

export default defineConfig({
  entry: {
    index: "src/runtime.ts",
    oxlint: "src/oxlint.ts",
    testing: "src/testing.ts",
  },
  format: ["esm"],
  platform: "neutral",
  dts: true,
  outDir: "dist",
  clean: true,
});
