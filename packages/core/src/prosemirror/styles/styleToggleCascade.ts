import type { TextFormatting } from "../../types/document";
import { mergeTextFormatting, STYLE_TOGGLE_KEYS } from "../../utils/textFormattingMerge";

type StyleToggleKey = (typeof STYLE_TOGGLE_KEYS)[number];

type ToggleState = {
  value: boolean;
  defaultsActive: boolean;
};

/** Internal formatting value whose toggle history survives copying and separate cascade passes. */
type StyleTextFormattingCascade = {
  formatting: TextFormatting | undefined;
  toggleStates: ReadonlyMap<StyleToggleKey, ToggleState>;
  type: "styleToggleCascade";
};

type StyleToggleLevel =
  | {
      formatting: TextFormatting | undefined;
      type: "defaults" | "style" | "direct";
    }
  | {
      cascade: StyleTextFormattingCascade;
      type: "carried";
    };

type StyleToggleCascadeOptions = {
  /** Existing ordinary-property merge whose toggle fields this cascade replaces. */
  ordinaryFormatting: TextFormatting | undefined;
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
): StyleTextFormattingCascade {
  let formatting = options ? mergeTextFormatting(undefined, options.ordinaryFormatting) : undefined;
  const mergeOrdinaryFormatting = options === undefined;
  const toggleStates = new Map<StyleToggleKey, ToggleState>();

  for (const level of levels) {
    if (level.type === "carried") {
      if (mergeOrdinaryFormatting) {
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
          value: inherited?.defaultsActive === true ? true : !(inherited?.value ?? false),
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
