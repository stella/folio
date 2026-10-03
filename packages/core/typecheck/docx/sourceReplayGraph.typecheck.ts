import type { Paragraph } from "../../src/types/document";
import type { SourceReplayGraph } from "../../src/docx/documentSource";

export const assertReadonlySourceGraph = (paragraph: SourceReplayGraph<Paragraph>) => {
  // @ts-expect-error Source-owned blocks cannot gain runs in place.
  paragraph.content.push({ type: "run", content: [] });
  // @ts-expect-error Source-owned properties cannot change in place.
  paragraph.paraId = "00000002";
  for (const child of paragraph.content) {
    if (child.type !== "run") continue;
    // @ts-expect-error Run descendants are immutable too.
    child.content.push({ type: "text", text: "lost edit" });
    for (const item of child.content) {
      if (item.type !== "text") continue;
      // @ts-expect-error Text descendants cannot change beneath an unchanged identity.
      item.text = "lost edit";
    }
  }
};
