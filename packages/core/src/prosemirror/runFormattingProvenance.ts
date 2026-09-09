import type { TextFormatting } from "../types/document";
import {
  RUN_FORMATTING_BOOLEAN_PROPERTIES,
  RUN_FORMATTING_PROPERTY_SPECS,
  RUN_FORMATTING_VALUE_PROPERTIES,
  type AuthoredRunFormattingValues,
  type RunFormattingBooleanProperty,
  type RunFormattingOverrideAttrs,
} from "./schema/marks";

const isRunFormattingProperty = (
  property: string,
): property is keyof typeof RUN_FORMATTING_PROPERTY_SPECS =>
  Object.hasOwn(RUN_FORMATTING_PROPERTY_SPECS, property);

export const hasAuthoredRunFormattingProvenance = ({
  _authoredOff,
  _authoredOn,
  _authoredValues,
}: RunFormattingOverrideAttrs): boolean =>
  Array.isArray(_authoredOn) ||
  Array.isArray(_authoredOff) ||
  (_authoredValues != null && Object.keys(_authoredValues).length > 0);

/** Whether an override mark still carries any rendering or authorship signal. */
export const hasRunFormattingOverrideAttrs = (attrs: RunFormattingOverrideAttrs): boolean => {
  if (hasAuthoredRunFormattingProvenance(attrs)) {
    return true;
  }
  return Object.values(attrs).some((value) => {
    if (Array.isArray(value)) {
      return value.length > 0;
    }
    if (typeof value === "object" && value !== null) {
      return Object.keys(value).length > 0;
    }
    return value !== null && value !== undefined;
  });
};

/** Reconstruct the last imported/reconciled direct baseline; inherited values never enter here. */
export const authoredRunFormattingFromAttrs = (
  attrs: RunFormattingOverrideAttrs,
): TextFormatting | undefined => {
  if (!hasAuthoredRunFormattingProvenance(attrs)) {
    return undefined;
  }

  const formatting: TextFormatting = {};
  for (const property of attrs._authoredOn ?? []) {
    Reflect.set(formatting, property, true);
  }
  for (const property of attrs._authoredOff ?? []) {
    Reflect.set(formatting, property, false);
  }
  for (const [property, value] of Object.entries(attrs._authoredValues ?? {})) {
    if (value !== undefined && isRunFormattingProperty(property)) {
      Reflect.set(formatting, property, value);
    }
  }
  return formatting;
};

/** Replace the direct baseline while preserving current rendering attributes. */
export const withAuthoredRunFormatting = (
  attrs: RunFormattingOverrideAttrs,
  formatting: TextFormatting | undefined,
): RunFormattingOverrideAttrs => {
  const authoredOn: RunFormattingBooleanProperty[] = [];
  const authoredOff: RunFormattingBooleanProperty[] = [];
  const authoredValues: AuthoredRunFormattingValues = {};

  for (const property of RUN_FORMATTING_BOOLEAN_PROPERTIES) {
    const value = formatting?.[property];
    if (value === true) {
      authoredOn.push(property);
    } else if (value === false) {
      authoredOff.push(property);
    }
  }
  for (const property of RUN_FORMATTING_VALUE_PROPERTIES) {
    const value = formatting?.[property];
    if (value !== undefined) {
      Reflect.set(authoredValues, property, value);
    }
  }

  const next = { ...attrs };
  delete next._authoredOn;
  delete next._authoredOff;
  delete next._authoredValues;
  if (authoredOn.length > 0) {
    next._authoredOn = authoredOn;
  }
  if (authoredOff.length > 0) {
    next._authoredOff = authoredOff;
  }
  if (Object.keys(authoredValues).length > 0) {
    next._authoredValues = authoredValues;
  }
  if (
    authoredOn.length === 0 &&
    authoredOff.length === 0 &&
    Object.keys(authoredValues).length === 0
  ) {
    // An empty array is the sparse known-empty discriminator. It appears only
    // on an override needed for rendering inherited/suppressed state; without
    // it, the legacy rendering attrs are indistinguishable from direct state.
    next._authoredOn = [];
  }
  return next;
};
