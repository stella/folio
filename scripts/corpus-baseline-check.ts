/**
 * Check the committed corpus files against themselves.
 *
 * Cheap enough for pull-request CI: it downloads nothing, runs no invariant and
 * reads only what the repository carries. What it cannot answer is whether the
 * recorded numbers still hold; that is the nightly's job, and it needs the
 * corpus to answer it.
 *
 *   bun scripts/corpus-baseline-check.ts
 */

import {
  corpusValidityIssues,
  loadCommittedCorpusFiles,
  renderValidityIssues,
} from "./lib/corpus-baseline-validity";

const issues = corpusValidityIssues(await loadCommittedCorpusFiles());
if (issues.length === 0) {
  process.stdout.write("corpus baselines are internally consistent\n");
} else {
  process.stderr.write(
    `${issues.length} inconsistency(ies) in the committed corpus files:\n${renderValidityIssues(issues)}\n`,
  );
  process.exitCode = 1;
}
