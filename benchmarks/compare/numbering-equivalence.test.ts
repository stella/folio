import { expect, test } from "bun:test";
import { compareDocx } from "@stll/folio-core";

import { buildDocumentPackage } from "./documents";
import { checkInvariants } from "./invariants";
import { equivalentNumberingAliases } from "./numbering-equivalence";
import { zipPackage, type DocxPackage } from "./package-xml";
import { applyVariant } from "./variants";

const OPTIONS = { author: "folio compare benchmark", timestamp: "2000-01-01T00:00:00.000Z" };

const listParts = () => buildDocumentPackage({ documentClass: "lists", size: "s" });

const withSecondInstance = (
  parts: DocxPackage,
  options: { remap: "all" | "first"; startOverride?: number; differentFormat?: boolean },
): DocxPackage => {
  const result = new Map(parts);
  const numbering = result.get("word/numbering.xml");
  const document = result.get("word/document.xml");
  if (typeof numbering !== "string" || typeof document !== "string") {
    throw new Error("The list fixture needs document and numbering XML");
  }
  const abstract = options.differentFormat
    ? numbering.match(/<w:abstractNum w:abstractNumId="0">[\s\S]*?<\/w:abstractNum>/u)?.at(0)
    : null;
  if (options.differentFormat && !abstract) throw new Error("Missing abstract numbering");
  const secondAbstract = abstract
    ?.replace('w:abstractNumId="0"', 'w:abstractNumId="1"')
    .replace('<w:numFmt w:val="decimal"/>', '<w:numFmt w:val="lowerRoman"/>');
  const instance =
    '<w:num w:numId="2"><w:abstractNumOverride w:val="0"/>' +
    `<w:abstractNumId w:val="${options.differentFormat ? "1" : "0"}"/>` +
    (options.startOverride === undefined
      ? ""
      : `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${String(options.startOverride)}"/></w:lvlOverride>`) +
    "</w:num>";
  result.set(
    "word/numbering.xml",
    numbering.replace("</w:numbering>", `${secondAbstract ?? ""}${instance}</w:numbering>`),
  );
  result.set(
    "word/document.xml",
    options.remap === "all"
      ? document.replaceAll('<w:numId w:val="1"/>', '<w:numId w:val="2"/>')
      : document.replace('<w:numId w:val="1"/>', '<w:numId w:val="2"/>'),
  );
  return result;
};

test("numbering alias equivalence preserves effective levels and sequence groups", async () => {
  const baseParts = listParts();
  const base = await zipPackage(baseParts);
  expect(
    await equivalentNumberingAliases(
      base,
      await zipPackage(withSecondInstance(baseParts, { remap: "all" })),
    ),
  ).toBe(true);
  expect(
    await equivalentNumberingAliases(
      base,
      await zipPackage(withSecondInstance(baseParts, { remap: "first" })),
    ),
  ).toBe(false);
  expect(
    await equivalentNumberingAliases(
      await zipPackage(withSecondInstance(baseParts, { remap: "first" })),
      base,
    ),
  ).toBe(false);
  expect(
    await equivalentNumberingAliases(
      base,
      await zipPackage(withSecondInstance(baseParts, { remap: "all", startOverride: 7 })),
    ),
  ).toBe(false);
  expect(
    await equivalentNumberingAliases(
      base,
      await zipPackage(withSecondInstance(baseParts, { remap: "all", differentFormat: true })),
    ),
  ).toBe(false);
});

test("round-trip gate accepts a rebased numId only when the list remains equivalent", async () => {
  const parts = listParts();
  const targetParts = applyVariant({ parts, variant: "numbering" });
  if (!targetParts) throw new Error("Numbering variant did not apply");
  const base = await zipPackage(parts);
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
    validate: null,
  });
  expect(
    result.outcomes.find(({ invariant }) => invariant === "accept-returns-target")?.status,
  ).toBe("passed");
  expect(result.outcomes.find(({ invariant }) => invariant === "reject-returns-base")?.status).toBe(
    "passed",
  );
});
