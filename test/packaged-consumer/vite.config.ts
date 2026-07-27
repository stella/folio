import { defineConfig } from "vite";

// Production build of a real downstream consumer against the PACKED folio
// tarballs. The point is bundler fidelity, not a runnable app: Vite/rolldown
// must resolve everything @stll/folio-react and @stll/folio-core reference,
// including worker entries constructed via `new URL(..., import.meta.url)`
// (the `vite:worker-import-meta-url` plugin turns each into its own build
// entry). A worker specifier pointing at a file the dist never emitted aborts
// this build with UNRESOLVED_ENTRY — exactly the failure a source-consuming
// build (the playground) cannot surface.
//
// Bundle the workspace packages under test so their internals, assets, and the
// worker inside folio-core go through the bundler. Externalize everything else
// (React, ProseMirror, yjs, jszip, …): a consumer installs those itself, and
// leaving them external keeps the build fast and focused on workspace output.
const BUNDLED_WORKSPACE_PACKAGES = [
  "@stll/docx-core",
  "@stll/folio-core",
  "@stll/folio-react",
] as const;

const isBundledWorkspacePackage = (id: string): boolean =>
  BUNDLED_WORKSPACE_PACKAGES.some(
    (packageName) => id === packageName || id.startsWith(`${packageName}/`),
  );

const isBundled = (id: string): boolean =>
  id.startsWith(".") || id.startsWith("/") || isBundledWorkspacePackage(id);

export default defineConfig({
  build: {
    outDir: "dist",
    lib: {
      entry: "src/main.ts",
      formats: ["es"],
      fileName: "main",
    },
    rollupOptions: {
      external: (id: string) => !isBundled(id),
    },
  },
});
