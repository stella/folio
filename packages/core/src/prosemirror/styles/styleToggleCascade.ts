import type { TextFormatting } from "../../types/document";
import type { DeepReadonly } from "../../types/documentModel";
import {
  type MergedTextFormatting,
  mergeTextFormatting,
  STYLE_TOGGLE_KEYS,
} from "../../utils/textFormattingMerge";

export type StyleToggleKey = (typeof STYLE_TOGGLE_KEYS)[number];

type ToggleState = {
  value: boolean;
  defaultsActive: boolean;
};

/**
 * Internal formatting value whose toggle history survives copying and separate
 * cascade passes. Its formatting shares nested values with the levels it read
 * (a style resolver's results among them), so it is read-only; TypeScript lets
 * a readonly literal match a writable parameter, so there is no writable
 * variant to fall into.
 */
type StyleTextFormattingCascade<
  Formatting extends DeepReadonly<TextFormatting> =
    DeepReadonly<TextFormatting>,
> = {
  formatting: Formatting | undefined;
  toggleStates: ReadonlyMap<StyleToggleKey, ToggleState>;
  type: "styleToggleCascade";
};

type StyleToggleLevel<
  Formatting extends DeepReadonly<TextFormatting> =
    DeepReadonly<TextFormatting>,
> =
  | {
      formatting: Formatting | undefined;
      type: "defaults" | "style" | "direct";
    }
  | {
      cascade: StyleTextFormattingCascade<Formatting>;
      type: "carried";
    };

type StyleToggleCascadeOptions<
  Formatting extends DeepReadonly<TextFormatting> =
    DeepReadonly<TextFormatting>,
> = {
  /** Existing ordinary-property merge whose toggle fields this cascade replaces. */
  ordinaryFormatting: Formatting | undefined;
};

type CascadeStyleFromResolvedBaseOptions = {
  baseFormatting: DeepReadonly<TextFormatting> | undefined;
  defaultsActiveProperties: readonly StyleToggleKey[] | undefined;
  styleFormatting: DeepReadonly<TextFormatting> | undefined;
};

/** Continue a style cascade from a compact resolved paragraph snapshot. */
export const cascadeStyleFromResolvedBase = ({
  baseFormatting,
  defaultsActiveProperties,
  styleFormatting,
}: CascadeStyleFromResolvedBaseOptions):
  DeepReadonly<TextFormatting> | undefined => {
  const defaultsActive = new Set(defaultsActiveProperties);
  const toggleStates = new Map<StyleToggleKey, ToggleState>();
  for (const key of STYLE_TOGGLE_KEYS) {
    const value = baseFormatting?.[key];
    if (value !== undefined) {
      toggleStates.set(key, { value, defaultsActive: defaultsActive.has(key) });
    }
  }
  return cascadeStyleTextFormatting([
    {
      cascade: {
        formatting: baseFormatting,
        toggleStates,
        type: "styleToggleCascade",
      },
      type: "carried",
    },
    { formatting: styleFormatting, type: "style" },
  ]).formatting;
};

/**
 * Resolve run formatting across OOXML style hierarchy levels.
 *
 * Ordinary properties retain Folio's existing low-to-high, last-defined merge. Toggle
 * properties reverse inherited state when a style level states `true`; `false` leaves inherited
 * style state unchanged, while direct formatting remains an absolute value. The returned value
 * carries its toggle history explicitly, so paragraph and character style resolution may happen
 * in separate passes without an object-identity contract.
 */
export function cascadeStyleTextFormatting(
  levels: readonly StyleToggleLevel[],
  options?: StyleToggleCascadeOptions,
): StyleTextFormattingCascade;
export function cascadeStyleTextFormatting(
  levels: readonly StyleToggleLevel[],
  options?: StyleToggleCascadeOptions,
): StyleTextFormattingCascade {
  // Every value assigned here is a merge with a source, a new record, so the
  // toggle writes below land in this cascade's own record, never in a level.
  let formatting: MergedTextFormatting | undefined = options?.ordinaryFormatting
    ? mergeTextFormatting(undefined, options.ordinaryFormatting)
    : undefined;
  const mergeOrdinaryFormatting = options === undefined;
  const toggleStates = new Map<StyleToggleKey, ToggleState>();

  for (const level of levels) {
    if (level.type === "carried") {
      if (mergeOrdinaryFormatting && level.cascade.formatting) {
        formatting = mergeTextFormatting(formatting, level.cascade.formatting);
      }
      for (const [key, state] of level.cascade.toggleStates) {
        toggleStates.set(key, state);
      }
      continue;
    }
    if (!level.formatting) {
      continue;
    }

    if (mergeOrdinaryFormatting) {
      formatting = mergeTextFormatting(formatting, level.formatting);
    }
    for (const key of STYLE_TOGGLE_KEYS) {
      const value = level.formatting[key];
      if (value === undefined) {
        continue;
      }

      const inherited = toggleStates.get(key);
      if (level.type === "direct") {
        toggleStates.set(key, { value, defaultsActive: false });
      } else if (!value) {
        continue;
      } else if (level.type === "defaults") {
        toggleStates.set(key, { value: true, defaultsActive: true });
      } else {
        toggleStates.set(key, {
          value:
            inherited?.defaultsActive === true
              ? true
              : !(inherited?.value ?? false),
          defaultsActive: inherited?.defaultsActive === true,
        });
      }
    }
  }

  if (!formatting && toggleStates.size > 0) {
    formatting = {};
  }
  for (const [key, state] of toggleStates) {
    if (formatting) {
      formatting[key] = state.value;
    }
  }
  return { formatting, toggleStates, type: "styleToggleCascade" };
}
