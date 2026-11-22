import type { TableWidthType } from "@stll/docx-core/model";

const TABLE_WIDTH_KIND = {
  auto: "automatic",
  nil: "unspecified",
  dxa: "twips",
  pct: "percentage",
} as const satisfies Record<TableWidthType, string>;

export const readTableWidthKind = (type: TableWidthType | undefined) =>
  type === undefined ? "unspecified" : TABLE_WIDTH_KIND[type];
