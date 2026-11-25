import { defineComponent, h } from "vue";

export const LocaleLoading = defineComponent({
  setup(_, { slots }) {
    return () =>
      h(
        "div",
        { class: "docx-editor-vue__loading", "aria-busy": "true" },
        slots.default?.() ?? [h("progress", { "aria-hidden": "true" })],
      );
  },
});
