import type { TableRow } from "../../model/document";
import type { DocumentIdLease } from "../../model/identity";

export const TEST_ID_LEASE = {
  paragraph: { start: 1, end: 0x7fffffff },
  revision: { start: 1, end: 0x7fffffff },
  comment: { start: 1, end: 0x7fffffff },
  table: { prefix: "test", start: 0, end: Number.MAX_SAFE_INTEGER },
  row: { prefix: "test", start: 0, end: Number.MAX_SAFE_INTEGER },
  cell: { prefix: "test", start: 0, end: Number.MAX_SAFE_INTEGER },
} satisfies DocumentIdLease;

const base32 = (value: number): string => {
  const alphabet = "abcdefghijklmnopqrstuvwxyz234567";
  const digits: string[] = [];
  do {
    digits.push(alphabet.charAt(value % 32));
    value = Math.floor(value / 32);
  } while (value > 0);
  return digits.reverse().join("");
};

export const testCellIds = (blockIds: readonly string[]): string[] =>
  blockIds.map((id) => `c_test_${base32(Number.parseInt(id, 16))}`);

/** Populate fixture identities before the fixture is submitted as an operation. */
export const identifyTestRow = (row: TableRow): TableRow => {
  const firstParagraph = (value: unknown): string | undefined => {
    if (value === null || typeof value !== "object") return undefined;
    if (Array.isArray(value)) {
      for (const entry of value) {
        const found = firstParagraph(entry);
        if (found !== undefined) return found;
      }
      return undefined;
    }
    if ("paraId" in value && typeof value.paraId === "string") return value.paraId;
    for (const entry of Object.values(value)) {
      const found = firstParagraph(entry);
      if (found !== undefined) return found;
    }
    return undefined;
  };
  let ordinal = 0;
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const entry of value) visit(entry);
      return;
    }
    if ("type" in value) {
      const paragraph = firstParagraph(value);
      const counter =
        (paragraph === undefined ? 0 : Number.parseInt(paragraph, 16)) * 1024 + ordinal;
      if (["table", "tableRow", "tableCell"].includes(String(value.type))) ordinal += 1;
      switch (value.type) {
        case "table":
          if (!("tableId" in value)) Object.assign(value, { tableId: `t_test_${base32(counter)}` });
          break;
        case "tableRow":
          if (!("rowId" in value)) Object.assign(value, { rowId: `r_test_${base32(counter)}` });
          break;
        case "tableCell":
          if (!("cellId" in value)) Object.assign(value, { cellId: `c_test_${base32(counter)}` });
          break;
      }
    }
    for (const entry of Object.values(value)) visit(entry);
  };
  visit(row);
  return row;
};
