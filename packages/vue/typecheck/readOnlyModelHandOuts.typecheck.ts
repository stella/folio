import type { DocxEditorRef } from "../src/components/DocxEditor/types";

// The paged editor handle's document is built over the editor's model and
// shares its parts, so it is handed out read-only; the editor ref's own
// `getDocument` hands the host a copy it may write.

export const pagedEditorDocument = (ref: DocxEditorRef): void => {
  const paged = ref.getEditorRef()?.getDocument();
  if (paged) {
    // @ts-expect-error the paged editor's document shares the model's parts
    paged.package.document.content.length = 0;
    // @ts-expect-error the paged editor's document shares the model's parts
    paged.package.headers?.clear();
  }
  const copy = ref.getDocument();
  if (copy) copy.package.document.content.length = 0;
};
