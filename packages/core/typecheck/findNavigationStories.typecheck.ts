import { OP_STORIES, type OpStory } from "@stll/docx-core/ops";

// The previous mount-sequence fixtures covered body/header only. Bind this
// matrix to the story union so a new story cannot omit virtual-page navigation.
export const navigationStories = {
  main: OP_STORIES.MAIN,
  header: { kind: "header", rId: "h-default" },
  footer: { kind: "footer", rId: "f-default" },
  footnote: { kind: "footnote", id: 1 },
  endnote: { kind: "endnote", id: 2 },
} as const satisfies Record<typeof OP_STORIES.MAIN | Exclude<OpStory, string>["kind"], OpStory>;
