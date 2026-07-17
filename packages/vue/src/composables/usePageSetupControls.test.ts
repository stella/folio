import { describe, test, expect, mock } from "bun:test";
import { ref } from "vue";
import type { Document } from "@stll/folio-core/types/document";
import { usePageSetupControls } from "./usePageSetupControls";

function makeDoc(): Document {
  return {
    package: {
      document: { content: [], finalSectionProperties: { marginLeft: 1440 } },
    },
  };
}

describe("usePageSetupControls.handlePageSetupApply — readOnly guard", () => {
  test("is a no-op when readOnly is true (mirrors React's DocxEditor.tsx guard)", () => {
    const doc = makeDoc();
    const onChange = mock((_doc: Document) => {});
    const reLayout = mock(() => {});
    const stateTick = ref(0);
    const { handlePageSetupApply } = usePageSetupControls({
      editorView: ref(null),
      getDocument: () => doc,
      readOnly: ref(true),
      stateTick,
      reLayout,
      onChange,
    });

    handlePageSetupApply({ marginLeft: 720 });

    expect(doc.package.document.finalSectionProperties?.marginLeft).toBe(1440);
    expect(onChange).not.toHaveBeenCalled();
    expect(reLayout).not.toHaveBeenCalled();
    expect(stateTick.value).toBe(0);
  });

  test("applies the change when not readOnly", () => {
    const doc = makeDoc();
    const onChange = mock((_doc: Document) => {});
    const reLayout = mock(() => {});
    const stateTick = ref(0);
    const { handlePageSetupApply } = usePageSetupControls({
      editorView: ref(null),
      getDocument: () => doc,
      readOnly: ref(false),
      stateTick,
      reLayout,
      onChange,
    });

    handlePageSetupApply({ marginLeft: 720 });

    expect(doc.package.document.finalSectionProperties?.marginLeft).toBe(720);
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(reLayout).toHaveBeenCalledTimes(1);
    expect(stateTick.value).toBe(1);
  });
});
