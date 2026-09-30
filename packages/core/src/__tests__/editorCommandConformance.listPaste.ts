import { DOCUMENT_SHAPES } from "./documentShapes";
import { EXTRA_OPERATIONS } from "./editorCommandConformance";
import { SELECTION_PLACEMENTS } from "./editorHarness";

// Direct-list shapes exercise numbering carried on paragraph marks. Style
// rebasing has separate gaps, so its shapes keep their existing coverage.
const shapes = DOCUMENT_SHAPES.filter(
  ({ features }) =>
    !features.includes("style-numbering") &&
    (features.includes("list-decimal") || features.includes("list-bullet")),
);
const operations = EXTRA_OPERATIONS.filter(
  ({ id }) => id === "paste:paragraphs" || id === "paste:copied-blocks",
);

// The same pairs drive the regression suite and gap exclusions.
const pairs = shapes.flatMap((shape) => operations.map((operation) => ({ shape, operation })));

export const LIST_PASTE_RESOLUTION_CASES = pairs.flatMap(({ shape, operation }) =>
  SELECTION_PLACEMENTS.map((placement) => ({ shape, operation, placement })),
);

export const LIST_PASTE_RESOLUTION_KEYS = pairs.map(({ shape, operation }) => ({
  shape: shape.id,
  operation: operation.id,
}));
