/**
 * Compile-time proof that a reserved-value map is total over its model type.
 *
 * The registry's guarantee is that a field added to the model without a
 * recorded decision fails `bun run typecheck`. Asserting that over a real model
 * type is impossible — a complete map compiles, and a proof needs the
 * incomplete case — so the claim is made against a fixture type here, where a
 * missing decision is deliberate and `@ts-expect-error` fails the build if the
 * gate ever stops firing.
 */

import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import { NO_RESERVED_VALUE, type ReservedValueDisposition } from "./disposition";

type FixtureFormatting = {
  alignment?: string;
  outlineLevel?: number;
};

const INCOMPLETE = {
  alignment: NO_RESERVED_VALUE,
  // @ts-expect-error `outlineLevel` has no decision, so the map is not total
} satisfies Record<keyof FixtureFormatting, ReservedValueDisposition>;

const COMPLETE = {
  alignment: NO_RESERVED_VALUE,
  outlineLevel: NO_RESERVED_VALUE,
} satisfies Record<keyof FixtureFormatting, ReservedValueDisposition>;

// A key the model type does not declare is rejected by the alias constraint.
// @ts-expect-error `notAField` is not a field of FixtureFormatting
type RejectsUnknownField = ExhaustiveFields<FixtureFormatting, "alignment" | "notAField">;

type Assert<Claim extends true> = Claim;

type CompleteResolvesToTheModelType = Assert<
  ExhaustiveFields<FixtureFormatting, keyof typeof COMPLETE> extends FixtureFormatting
    ? true
    : false
>;

type IncompleteResolvesToNever = Assert<
  [ExhaustiveFields<FixtureFormatting, keyof typeof INCOMPLETE>] extends [never] ? true : false
>;

export type ReservedValueTotalityProof = [
  CompleteResolvesToTheModelType,
  IncompleteResolvesToNever,
  RejectsUnknownField,
];
