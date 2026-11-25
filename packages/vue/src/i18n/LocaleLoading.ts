import { h } from "vue";

export const LocaleLoading = () =>
  h("div", { class: "docx-editor-vue__loading", "aria-busy": "true" }, [
    h("progress", { "aria-hidden": "true" }),
  ]);
