import { defineConfig } from "tsdown";

export default defineConfig({
  entry: [
    "src/index.ts",
    "src/model/document.ts",
    "src/ops/documentOps.ts",
    "src/projection.ts",
    "src/schema/sequenceChildren.ts",
    "src/xml/streamingXmlParser.ts",
    "src/xml/xmlNamespaceContext.ts",
    "src/xml/xmlResourceLimits.ts",
    "src/zip/archive.ts",
    "src/zip/stream.ts",
  ],
  format: ["esm"],
  platform: "neutral",
  dts: { generator: "oxc" },
  outDir: "dist",
  clean: true,
  copy: [{ from: "src/generated/docx_kernel_bg.wasm", to: "dist" }],
});
