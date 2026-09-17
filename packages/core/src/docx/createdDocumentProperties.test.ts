import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createEmptyDocument } from "../utils/createDocument";

import { createDocx } from "./rezip";

const CORE_PROPERTIES_PATH = "docProps/core.xml";
const EXTENDED_PROPERTIES_PATH = "docProps/app.xml";

/** `XX.YYYY`: the only form the extended-properties `AppVersion` may take. */
const APP_VERSION_SCHEMA_FORM = /^\d{1,2}\.\d{4}$/u;

const partsOf = async (
  buffer: ArrayBuffer,
): Promise<{ coreProperties: string; extendedProperties: string }> => {
  const zip = await JSZip.loadAsync(buffer);
  const [coreProperties, extendedProperties] = await Promise.all([
    zip.file(CORE_PROPERTIES_PATH)?.async("text"),
    zip.file(EXTENDED_PROPERTIES_PATH)?.async("text"),
  ]);
  expect(coreProperties).toBeString();
  expect(extendedProperties).toBeString();
  return { coreProperties: coreProperties ?? "", extendedProperties: extendedProperties ?? "" };
};

const elementValue = (xml: string, name: string): string | undefined =>
  new RegExp(
    `<(?:[^\\s<>/:]+:)?${name}(?:\\s[^<>]*)?>([^<]*)</(?:[^\\s<>/:]+:)?${name}>`,
    "u",
  ).exec(xml)?.[1];

describe("the document properties a newly created package states", () => {
  test("writes no author or application by default", async () => {
    const { coreProperties, extendedProperties } = await partsOf(
      await createDocx(createEmptyDocument()),
    );

    expect(elementValue(coreProperties, "creator")).toBeUndefined();
    expect(elementValue(extendedProperties, "Application")).toBeUndefined();
    expect(elementValue(extendedProperties, "AppVersion")).toBeUndefined();
  });

  test("writes the host-supplied creator and application", async () => {
    const { coreProperties, extendedProperties } = await partsOf(
      await createDocx(createEmptyDocument(), {
        creator: "Marie Nováková",
        application: "Example Host 3.2",
      }),
    );

    expect(elementValue(coreProperties, "creator")).toBe("Marie Nováková");
    expect(elementValue(extendedProperties, "Application")).toBe("Example Host 3.2");
  });

  test("states an application version of the schema form alongside an application", async () => {
    const { extendedProperties } = await partsOf(
      await createDocx(createEmptyDocument(), { application: "Example Host" }),
    );

    expect(elementValue(extendedProperties, "AppVersion")).toMatch(APP_VERSION_SCHEMA_FORM);
  });

  test("escapes a supplied value rather than letting it close its element", async () => {
    const { coreProperties, extendedProperties } = await partsOf(
      await createDocx(createEmptyDocument(), {
        creator: "</dc:creator><dc:title>injected</dc:title><dc:creator>",
        application: "Smith & Co <Legal>",
      }),
    );

    expect(elementValue(coreProperties, "title")).toBeUndefined();
    expect(elementValue(extendedProperties, "Application")).toBe("Smith &amp; Co &lt;Legal&gt;");
  });
});

describe("the document properties a package carrying a source states", () => {
  test("keeps the properties of the source package over the supplied ones", async () => {
    const source = await createDocx(createEmptyDocument(), {
      creator: "Source Author",
      application: "Source Application",
    });

    const { coreProperties, extendedProperties } = await partsOf(
      await createDocx(
        { ...createEmptyDocument(), originalBuffer: source },
        { creator: "Other Author", application: "Other Application" },
      ),
    );

    expect(elementValue(coreProperties, "creator")).toBe("Source Author");
    expect(elementValue(extendedProperties, "Application")).toBe("Source Application");
  });
});
