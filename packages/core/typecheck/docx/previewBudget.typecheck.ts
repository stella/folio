import { createPackagePreviewBudget } from "../../src/docx/previewBudget";

const budget = createPackagePreviewBudget();
const frame = { size: { width: 1, height: 1 }, wrap: { type: "inline" } } as const;

// A descriptor-backed preview has no image data and therefore no character allowance.
// @ts-expect-error `smartArt` is deliberately absent from PreviewBudgetOverrides.
budget.enforce({ smartArt: 0 });

// Nor can the ledger build one: its factory draws source-backed kinds only.
// @ts-expect-error `smartArt` is not a kind the ledger stamps into `src`.
budget.ledger.svgImage("smartArt", "<svg/>", frame);

// A preview names no relationship, so its frame cannot carry one.
// @ts-expect-error `rId` is not part of a preview's frame.
budget.ledger.svgImage("vmlShape", "<svg/>", { ...frame, rId: "rId1" });
