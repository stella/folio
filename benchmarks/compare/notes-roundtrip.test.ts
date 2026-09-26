import { expect, test } from "bun:test";
import { compareDocx } from "@stll/folio-core";

import { buildDocumentPackage } from "./documents";
import { checkInvariants } from "./invariants";
import { zipPackage } from "./package-xml";
import { applyVariant } from "./variants";

const OPTIONS = { author: "folio compare benchmark", timestamp: "2000-01-01T00:00:00.000Z" };

test.each(["notes", "everywhere"] as const)(
  "generated notes survive %s accept and reject round trips",
  async (variant) => {
    const parts = buildDocumentPackage({ documentClass: "notes", size: "s" });
    const footnotes = parts.get("word/footnotes.xml");
    const endnotes = parts.get("word/endnotes.xml");
    expect(footnotes).toContain("<w:footnoteRef/>");
    expect(endnotes).toContain("<w:endnoteRef/>");
    const targetParts = applyVariant({ parts, variant });
    if (!targetParts) throw new Error(`Notes fixture does not support ${variant}`);
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
    expect(result.outcomes.filter(({ status }) => status === "failed")).toEqual([]);
  },
);
