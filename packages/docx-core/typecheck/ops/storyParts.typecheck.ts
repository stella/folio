import type { StoryParts } from "../../src/ops/types";

const narrow = {
  sections: [{ index: 0, properties: {} }],
} satisfies StoryParts;
const complete = { sectionView: { type: "omitted" } } satisfies StoryParts;

// @ts-expect-error A complete section view and narrow section metadata have distinct ownership.
const overlapping = { ...narrow, ...complete } satisfies StoryParts;

void overlapping;
