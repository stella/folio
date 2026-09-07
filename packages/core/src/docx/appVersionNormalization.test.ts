import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import {
  appVersionInSchemaForm,
  normalizeAppVersionInExtendedProperties,
} from "./appVersionNormalization";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx, updateMultipleFiles } from "./rezip";

const SCHEMA_FORM = /^\d{1,2}\.\d{4}$/u;

const EXTENDED_PROPERTIES_PATH = "docProps/app.xml";

const extendedProperties = (appVersionElement: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
  `<Pages>1</Pages><Words>2</Words><Characters>11</Characters>` +
  appVersionElement +
  `<Lines>1</Lines><Paragraphs>1</Paragraphs>` +
  `</Properties>`;

describe("appVersionInSchemaForm", () => {
  test.each([
    ["1.0.0", "1.0000"],
    ["16.0000", "16.0000"],
    ["1.0", "1.0000"],
    ["abc", "1.0000"],
    ["", "1.0000"],
    ["12.1.3.4", "12.0000"],
    ["1234.5", "1.0000"],
  ])("%p becomes %p", (value, expected) => {
    expect(appVersionInSchemaForm(value)).toBe(expected);
  });
});

describe("normalizeAppVersionInExtendedProperties", () => {
  test("rewrites a three-part version and leaves the rest of the part alone", () => {
    const xml = extendedProperties("<AppVersion>1.0.0</AppVersion>");

    expect(normalizeAppVersionInExtendedProperties(xml)).toBe(
      extendedProperties("<AppVersion>1.0000</AppVersion>"),
    );
  });

  test("returns a part whose version already fits unchanged", () => {
    const xml = extendedProperties("<AppVersion>16.0000</AppVersion>");

    expect(normalizeAppVersionInExtendedProperties(xml)).toBe(xml);
  });

  test("returns a part that states no version unchanged", () => {
    const xml = extendedProperties("");

    expect(normalizeAppVersionInExtendedProperties(xml)).toBe(xml);
  });

  test("writes a value of the schema form for any version a part can state", () => {
    fc.assert(
      fc.property(fc.string(), (value) => {
        const normalized = normalizeAppVersionInExtendedProperties(
          extendedProperties(`<AppVersion>${value.replaceAll(/[<&]/gu, "")}</AppVersion>`),
        );
        const written = /<AppVersion>([^<]*)<\/AppVersion>/u.exec(normalized)?.[1] ?? "";

        expect(written).toMatch(SCHEMA_FORM);
        expect(normalizeAppVersionInExtendedProperties(normalized)).toBe(normalized);
      }),
      propertyConfig({ numRuns: 500 }),
    );
  });
});

const appVersionOf = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const xml = (await zip.file(EXTENDED_PROPERTIES_PATH)?.async("text")) ?? "";
  return /<AppVersion>([^<]*)<\/AppVersion>/u.exec(xml)?.[1] ?? "";
};

const extendedPropertiesOf = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  return (await zip.file(EXTENDED_PROPERTIES_PATH)?.async("text")) ?? "";
};

/** A package that states a version the schema form rejects. */
const packageStatingAppVersion = async (value: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(EXTENDED_PROPERTIES_PATH, extendedProperties(`<AppVersion>${value}</AppVersion>`));
  return await zip.generateAsync({ type: "arraybuffer" });
};

describe("a saved package states a version of the schema form", () => {
  test("the full repack rewrites it and changes nothing else in the part", async () => {
    const source = await packageStatingAppVersion("1.0.0");

    const saved = await repackDocx(await parseDocx(source));

    expect(await appVersionOf(saved)).toBe("1.0000");
    expect(await extendedPropertiesOf(saved)).toBe(
      (await extendedPropertiesOf(source)).replace("1.0.0", "1.0000"),
    );
  });

  test("the selective save rewrites it and changes nothing else in the part", async () => {
    const source = await packageStatingAppVersion("1.0.0");

    const saved = await updateMultipleFiles(source, new Map());

    expect(await appVersionOf(saved)).toBe("1.0000");
    expect(await extendedPropertiesOf(saved)).toBe(
      (await extendedPropertiesOf(source)).replace("1.0.0", "1.0000"),
    );
  });

  test("a package that already fits keeps the version it stated", async () => {
    const source = await packageStatingAppVersion("16.0000");

    expect(await appVersionOf(await repackDocx(await parseDocx(source)))).toBe("16.0000");
    expect(await appVersionOf(await updateMultipleFiles(source, new Map()))).toBe("16.0000");
  });

  test("a package that carries no extended properties gains none", async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx());
    zip.remove(EXTENDED_PROPERTIES_PATH);
    const source = await zip.generateAsync({ type: "arraybuffer" });

    const saved = await updateMultipleFiles(source, new Map());

    expect((await JSZip.loadAsync(saved)).file(EXTENDED_PROPERTIES_PATH)).toBeNull();
  });
});
