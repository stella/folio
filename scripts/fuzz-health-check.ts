import { appendFileSync, existsSync, readFileSync } from "node:fs";

import { HEALTH_MARKER } from "../test/consumer-scenarios/support/fuzz-health";

type HealthOptions = { log: string; outcome: string };

/** Missing or malformed reports fail closed, even after an earlier finding. */
export const checkFuzzHealth = ({ log, outcome }: HealthOptions) => {
  let completed = 0;
  let findings = 0;
  let pending = 0;
  const problems: string[] = [];
  for (const line of log.split("\n")) {
    const index = line.indexOf(HEALTH_MARKER);
    if (index === -1) continue;
    let report: unknown;
    try {
      report = JSON.parse(line.slice(index + HEALTH_MARKER.length));
    } catch {
      problems.push("Malformed fuzz health report");
      continue;
    }
    if (
      typeof report !== "object" ||
      report === null ||
      !("status" in report) ||
      !("completed" in report) ||
      typeof report.completed !== "number" ||
      !Number.isSafeInteger(report.completed) ||
      report.completed < 0
    ) {
      problems.push("Invalid fuzz health report");
      continue;
    }
    if (report.status === "started") {
      pending += 1;
      continue;
    }
    if (pending > 0) pending -= 1;
    completed += report.completed;
    if (report.status === "passed") continue;
    if (
      !("detail" in report) ||
      typeof report.detail !== "string" ||
      !report.detail.trim() ||
      report.detail.trim() === "undefined" ||
      report.detail.trim() === "null"
    ) {
      problems.push("Missing failure detail");
      continue;
    }
    if (report.status === "finding") findings += 1;
    else if (report.status === "infrastructure") problems.push(report.detail);
    else problems.push("Unknown fuzz health status");
  }
  if (pending > 0) problems.push(`${pending} fuzz runs started without a completion report`);
  if (completed === 0) problems.push("Zero completed cases or missing completion report");
  if (outcome !== "success" && (outcome !== "failure" || findings === 0)) {
    problems.push(`Fuzz step ${outcome}: crash, setup failure, or timeout without a finding`);
  }
  const result = { completed, findings, problems };
  if (problems.length > 0) return { status: "infrastructure", ...result };
  return {
    status: findings > 0 ? "finding" : "passed",
    completed,
    findings,
    problems,
  };
};

if (import.meta.main) {
  const logPath = process.argv.at(2);
  const outcome = process.env["FUZZ_STEP_OUTCOME"] ?? "unknown";
  const log = logPath && existsSync(logPath) ? readFileSync(logPath, "utf8") : "";
  const health = checkFuzzHealth({ log, outcome });
  const summary = `Fuzz health: **${health.status}**; ${health.completed} completed cases, ${health.findings} findings.\n${health.problems.map((problem) => `- ${problem}`).join("\n")}\n`;
  console.log(summary);
  const summaryPath = process.env["GITHUB_STEP_SUMMARY"];
  if (summaryPath) appendFileSync(summaryPath, summary);
  if (health.status === "infrastructure") {
    console.error("::error::Fuzz infrastructure failure; inspect the fuzz health summary and log");
    process.exitCode = 1;
  }
}
