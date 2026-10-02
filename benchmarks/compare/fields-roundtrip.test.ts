import { beforeAll, expect, test } from "bun:test";
import { compareDocx } from "@stll/folio-core";

import { buildDocumentPackage } from "./documents";
import { checkInvariants } from "./invariants";
import { zipPackage } from "./package-xml";
import { resolvePackageValidator, type PackageValidator } from "./validator";
import { applyVariant } from "./variants";

const OPTIONS = { author: "folio compare benchmark", timestamp: "2000-01-01T00:00:00.000Z" };
const SIMPLE_FIELD =
  /<w:fldSimple w:instr=" PAGE "><w:r>(<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>)<\/w:r><\/w:fldSimple>/gu;

let validate: PackageValidator | null = null;
beforeAll(async () => {
  validate = resolvePackageValidator();
  if (validate === null) return;
  // Validate the shared input while initializing the external SDK. Individual
  // cases still validate their own redlines and keep their existing timeout.
  const input = await zipPackage(buildDocumentPackage({ documentClass: "fields", size: "s" }));
  expect(validate(input)).toEqual([]);
});

test.each([
  ["ordinary", ""],
  ["locked and dirty", ' w:fldLock="1" w:dirty="1"'],
] as const)(
  "tracked %s simple fields survive schema and review round trips",
  async (_name, attrs) => {
    const baseParts = new Map(buildDocumentPackage({ documentClass: "fields", size: "s" }));
    const targetSource = applyVariant({ parts: baseParts, variant: "light" });
    if (!targetSource) throw new Error("Fields fixture does not support light");
    const targetParts = new Map(targetSource);
    for (const parts of [baseParts, targetParts]) {
      const document = parts.get("word/document.xml");
      if (typeof document !== "string") throw new Error("Missing document part");
      let replaced = 0;
      parts.set(
        "word/document.xml",
        document.replaceAll(SIMPLE_FIELD, (_field, resultXml: string) => {
          replaced++;
          return `<w:fldSimple w:instr=" PAGE "${attrs}><w:r>${resultXml}</w:r></w:fldSimple>`;
        }),
      );
      expect(replaced).toBeGreaterThan(0);
    }
    const base = await zipPackage(baseParts);
    const target = await zipPackage(targetParts);
    const compared = await compareDocx(base, target, OPTIONS);
    if (compared.isErr()) throw compared.error;
    const result = await checkInvariants({
      base,
      target,
      redlined: compared.value.buffer,
      changes: compared.value.changes,
      unsupported: compared.value.unsupported.map(({ reason }) => reason),
      expectation: "different",
      options: OPTIONS,
      validate,
    });
    expect(result.outcomes.filter(({ status }) => status === "failed")).toEqual([]);
  },
);

test.each([
  [
    "hyperlink",
    (resultXml: string) => `<w:hyperlink w:anchor="_page"><w:r>${resultXml}</w:r></w:hyperlink>`,
  ],
  ["wrapper", (resultXml: string) => `<w:dir w:val="ltr"><w:r>${resultXml}</w:r></w:dir>`],
  [
    "nested field",
    (resultXml: string) =>
      `<w:fldSimple w:instr=" NUMPAGES "><w:r>${resultXml}</w:r></w:fldSimple>`,
  ],
] as const)("refuses a tracked simple field with %s result content", async (_name, resultOf) => {
  const baseParts = new Map(buildDocumentPackage({ documentClass: "fields", size: "s" }));
  const targetSource = applyVariant({ parts: baseParts, variant: "light" });
  if (!targetSource) throw new Error("Fields fixture does not support light");
  const targetParts = new Map(targetSource);
  for (const parts of [baseParts, targetParts]) {
    const document = parts.get("word/document.xml");
    if (typeof document !== "string") throw new Error("Missing document part");
    parts.set(
      "word/document.xml",
      document.replaceAll(
        SIMPLE_FIELD,
        (_field, resultXml: string) =>
          `<w:fldSimple w:instr=" PAGE ">${resultOf(resultXml)}</w:fldSimple>`,
      ),
    );
  }
  const base = await zipPackage(baseParts);
  const target = await zipPackage(targetParts);
  expect(validate?.(base) ?? []).toEqual([]);
  expect(validate?.(target) ?? []).toEqual([]);
  const compared = await compareDocx(base, target, OPTIONS);
  expect(compared.isErr() && compared.error._tag).toBe("CompareDocxRoundTripError");
});
