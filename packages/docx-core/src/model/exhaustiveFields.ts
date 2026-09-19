/**
 * Compile-time totality over every field of a model type.
 *
 * `ExhaustiveFields<Source, Classified>` resolves to `Source` while
 * `Classified` names every key of `Source`, and to `never` otherwise. Alias it
 * next to a table that must decide something per field, and a field added to
 * the model without a decision turns the alias into `never`, so the module
 * stops compiling instead of silently dropping the field.
 *
 * ```ts
 * type ClassifiedBorderField = "style" | "color" | "size";
 * type ExhaustiveBorder = ExhaustiveFields<BorderSpec, ClassifiedBorderField>;
 * ```
 */
export type ExhaustiveFields<Source, Classified extends keyof Source> =
  Exclude<keyof Source, Classified> extends never ? Source : never;
