/**
 * Paragraph-style options for the Vue toolbar's style picker. Mirrors React's
 * StylePicker via the shared adapter helper: when the loaded document provides
 * paragraph styles, show those (real names + document order); otherwise fall
 * back to the built-in presets. The filter/sort + preview CSS live in the
 * adapter's `utils/stylePreview` (resolveParagraphStyleOptions /
 * getStylePreviewProps); only the i18n label resolution stays call-site.
 * Extracted from Toolbar.vue to keep it lean.
 */
import { computed, type ComputedRef } from "vue";
import {
  getStylePreviewProps,
  resolveParagraphStyleOptions,
  type StylePreviewProps,
} from "../utils/stylePreview";
import type { Style } from "@stll/folio-core/types/document";
import { paragraphStyles, type TranslationKey } from "../components/Toolbar/presets";

export type ResolvedStyle = {
  id: string;
  label: string;
  previewStyle: StylePreviewProps;
}

export type UseParagraphStyleOptionsReturn = {
  resolvedParagraphStyles: ComputedRef<ResolvedStyle[]>;
  currentStyleLabel: ComputedRef<string>;
}

export function useParagraphStyleOptions(opts: {
  documentStyles: () => Style[] | undefined;
  currentStyleId: () => string;
  t: (key: TranslationKey) => string;
}): UseParagraphStyleOptionsReturn {
  const resolvedParagraphStyles = computed<ResolvedStyle[]>(() => {
    const options = resolveParagraphStyleOptions(opts.documentStyles());
    if (options.length > 0) {
      return options.map((o) => {
        const preset = paragraphStyles.find((p) => p.id === o.styleId);
        return {
          id: o.styleId,
          label: preset ? opts.t(preset.nameKey) : o.name,
          previewStyle: getStylePreviewProps(o),
        };
      });
    }
    return paragraphStyles.map((p) => ({
      id: p.id,
      label: opts.t(p.nameKey),
      previewStyle: getStylePreviewProps({ styleId: p.id }),
    }));
  });

  const currentStyleLabel = computed(() => {
    const id = opts.currentStyleId() || "Normal";
    const s = resolvedParagraphStyles.value.find((ps) => ps.id === id);
    return s ? s.label : id;
  });

  return { resolvedParagraphStyles, currentStyleLabel };
}
