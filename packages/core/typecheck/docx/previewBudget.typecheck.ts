import { enforcePackagePreviewBudget } from "../../src/docx/previewBudget";

const image = { type: "image", preview: { kind: "diagram" } };

// A descriptor-backed preview has no image data and therefore no character allowance.
// @ts-expect-error `smartArt` is deliberately absent from PreviewBudgetOverrides.
enforcePackagePreviewBudget({ image }, { smartArt: 0 });
