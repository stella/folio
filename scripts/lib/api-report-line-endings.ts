import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

/** Check every report on disk, including obsolete exports and package folders. */
export const apiReportsWithCarriageReturns = (directory: string): string[] => {
  const invalid: string[] = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const full = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      for (const nested of apiReportsWithCarriageReturns(full)) {
        invalid.push(path.join(entry.name, nested));
      }
      continue;
    }
    if (entry.isFile() && entry.name.endsWith(".api.md") && readFileSync(full).includes(13)) {
      invalid.push(entry.name);
    }
  }
  return invalid.toSorted();
};
