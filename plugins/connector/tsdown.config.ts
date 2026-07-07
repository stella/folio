import { defineConfig } from "tsdown";

const entry = ["src/**/*.ts", "!src/**/*.test.ts", "!src/**/__tests__/**"];

const shared = {
  entry,
  format: ["esm"] as const,
  platform: "neutral" as const,
  outDir: "dist",
  unbundle: true,
};

export default defineConfig([
  { ...shared, dts: false, clean: false },
  { ...shared, dts: { emitDtsOnly: true }, clean: false },
]);
