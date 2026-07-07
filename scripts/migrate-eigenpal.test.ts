import { describe, expect, test } from "bun:test";

import { transformSource } from "./migrate-eigenpal";

describe("migrate-eigenpal codemod", () => {
  test("rewrites safe React, Vue, Nuxt, and stylesheet imports", () => {
    const input = `
import { DocxEditor } from "@eigenpal/docx-editor-react";
import "@eigenpal/docx-editor-react/styles.css";
import { en } from "@eigenpal/docx-editor-i18n";
import { DocxEditor as VueDocxEditor } from "@eigenpal/docx-editor-vue";
import "@eigenpal/docx-editor-vue/styles.css";
export { default } from "@eigenpal/nuxt-docx-editor";
`;

    const result = transformSource("src/app.tsx", input);

    expect(result.changed).toBe(true);
    expect(result.findings).toEqual([]);
    expect(result.text).toContain('from "@stll/folio-react/compat/eigenpal"');
    expect(result.text).toContain('"@stll/folio-react/standalone.css"');
    expect(result.text).toContain('from "@stll/folio-react/compat/eigenpal"');
    expect(result.text).toContain('from "@stll/folio-vue"');
    expect(result.text).toContain('"@stll/folio-vue/editor.css"');
    expect(result.text).toContain('from "@stll/folio-nuxt"');
  });

  test("rewrites safe core subpaths", () => {
    const input = `
import { setGoogleFontsEnabled } from "@eigenpal/docx-editor-core";
import type { Document } from "@eigenpal/docx-editor-core/types/document";
import { schema } from "@eigenpal/docx-editor-core/prosemirror/schema";
import { toProseDoc } from "@eigenpal/docx-editor-core/prosemirror/conversion";
`;

    const result = transformSource("src/core.ts", input);

    expect(result.changed).toBe(true);
    expect(result.findings).toEqual([]);
    expect(result.text).toContain('from "@stll/folio-core/compat/eigenpal"');
    expect(result.text).toContain('from "@stll/folio-core/types/document"');
    expect(result.text).toContain('from "@stll/folio-core/prosemirror/schema"');
    expect(result.text).toContain('from "@stll/folio-core/prosemirror/conversion"');
  });

  test("leaves incompatible subpaths untouched and reports them", () => {
    const input = `
import { PluginHost } from "@eigenpal/docx-editor-react/plugin-api";
import { DocxReviewer } from "@eigenpal/docx-editor-agents";
import { parseDocx } from "@eigenpal/docx-editor-core/docx";
import en from "@eigenpal/docx-editor-i18n/en";
`;

    const result = transformSource("src/integration.ts", input);

    expect(result.changed).toBe(false);
    expect(result.text).toBe(input);
    expect(result.findings.map((finding) => finding.line)).toEqual([2, 3, 4, 5]);
    expect(result.findings.map((finding) => finding.message)).toEqual([
      expect.stringContaining("React subpath has no direct folio equivalent"),
      expect.stringContaining("Agent API changed"),
      expect.stringContaining("Core public API is not drop-in"),
      expect.stringContaining("Locale subpaths cannot be rewritten safely"),
    ]);
  });

  test("rewrites package.json dependency names to installable folio ranges", () => {
    const input = `${JSON.stringify(
      {
        dependencies: {
          "@eigenpal/docx-editor-react": "^1.9.0",
          "@eigenpal/docx-editor-core": "^1.9.0",
          "@eigenpal/docx-editor-i18n": "^1.9.0",
        },
        devDependencies: {
          "@eigenpal/nuxt-docx-editor": "^1.9.0",
        },
      },
      null,
      2,
    )}\n`;

    const result = transformSource("package.json", input);
    const parsed = JSON.parse(result.text);

    expect(result.changed).toBe(true);
    expect(parsed.dependencies["@stll/folio-react"]).toBe("latest");
    expect(parsed.dependencies["@stll/folio-core"]).toBe("latest");
    expect(parsed.devDependencies["@stll/folio-nuxt"]).toBe("latest");
    expect(parsed.dependencies["@eigenpal/docx-editor-react"]).toBeUndefined();
    expect(parsed.dependencies["@eigenpal/docx-editor-i18n"]).toBeUndefined();
    expect(result.findings).toEqual([]);
  });

  test("rewrites Eigenpal specifiers in non-package JSON config files", () => {
    const input = `${JSON.stringify(
      {
        compilerOptions: {
          paths: {
            "@eigenpal/docx-editor-react": ["./vendor/docx-editor-react"],
          },
        },
      },
      null,
      2,
    )}\n`;

    const result = transformSource("tsconfig.json", input);

    expect(result.changed).toBe(true);
    expect(result.text).toContain("@stll/folio-react/compat/eigenpal");
  });
});
