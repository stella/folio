import { defineConfig } from "tsdown";

export default defineConfig({
  entry: ["src/index.ts", "src/model/document.ts", "src/projection.ts"],
  format: ["esm"],
  platform: "neutral",
  dts: true,
  outDir: "dist",
  clean: true,
  copy: [{ from: "src/generated/docx_kernel_bg.wasm", to: "dist" }],
});
