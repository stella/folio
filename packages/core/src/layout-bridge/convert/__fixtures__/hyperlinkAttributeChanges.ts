// A field-totality gate in a test file is read by no compiler, so the changed
// value per hyperlink attribute lives here, where the package typecheck
// compiles it.
import type { RecordField } from "../../../types/document";
import type { HyperlinkAttrs } from "../../../prosemirror/schema/marks";

/** One changed value per payload attribute of a hyperlink mark. */
export const HYPERLINK_ATTRIBUTE_CHANGES: {
  [Key in Exclude<
    RecordField<HyperlinkAttrs>,
    "_docxHyperlinkIndex"
  >]-?: Pick<HyperlinkAttrs, Key>;
} = {
  href: { href: "https://example.test/changed" },
  anchor: { anchor: "ChangedBookmark" },
  tooltip: { tooltip: "changed" },
  rId: { rId: "rId2" },
  target: { target: "_blank" },
  history: { history: false },
  docLocation: { docLocation: "changed" },
};
