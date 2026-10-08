#!/usr/bin/env bun
import { resolve } from "node:path";
import { reportsWithCarriageReturns } from "./lib/api-report-line-endings";

const reports = reportsWithCarriageReturns(resolve(import.meta.dir, "../api-reports"));
if (reports.length > 0) {
  console.error("API reports must use LF line endings:");
  for (const report of reports) console.error(`  api-reports/${report}`);
  process.exit(1);
}
