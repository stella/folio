/**
 * Compile-time proof that an attr patch cannot carry a raw model value.
 *
 * The claim a patch producer makes is that its `numPr` went through
 * `paragraphNumberingAttr`. While a producer returned `Record<string,
 * unknown>` the claim was unenforced: a `ParagraphFormatting.numPr` read off
 * a model object assigned straight into the patch, and `toProseDoc` stored
 * the model's union verbatim in an attr. A `ParagraphAttrsPatch` return type
 * makes that a type error, and `@ts-expect-error` fails `bun run typecheck`
 * if it ever stops being one.
 *
 * It lives beside the producers rather than in their test: a package's
 * `typecheck` runs over `tsconfig.build.json`, which excludes `*.test.ts`, so
 * an assertion in a test file is checked by nothing.
 */

import type { ParagraphNumberingOverride } from "@stll/docx-core/model";

import { paragraphNumberingAttr } from "../numberingAttr";
import type { ParagraphAttrsPatch } from "../schema/nodes";

declare const stated: ParagraphNumberingOverride;

const RAW_MODEL_VALUE: ParagraphAttrsPatch = {
  // @ts-expect-error the model's own union is not a minted attr value
  numPr: stated,
  // @ts-expect-error and neither is a record assembled by hand
  numPrFromStyle: { kind: "reference", numId: 1, ilvl: 0 },
};

const MINTED: ParagraphAttrsPatch = { numPr: paragraphNumberingAttr(stated) };

// A key the node spec does not declare is rejected outright, so a patch cannot
// smuggle a field past every other projection of the attrs.
const UNKNOWN_ATTR: ParagraphAttrsPatch = {
  // @ts-expect-error `numbering` is not a paragraph attr
  numbering: stated,
};

export type ParagraphAttrsPatchProof = [typeof RAW_MODEL_VALUE, typeof MINTED, typeof UNKNOWN_ATTR];
