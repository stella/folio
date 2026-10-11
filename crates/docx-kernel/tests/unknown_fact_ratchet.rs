#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::panic,
    clippy::unwrap_used
)]
// Unknown-fact ratchet over the repository's committed DOCX packages.
//
// Count typed unknown facts per (view, family, reason) across every committed
// package and compare with `unknown-fact-baseline.tsv`. Each fixture-specific
// unknown must match `unknown-fact-legitimate.tsv`, including during baseline
// updates. A count may only fall: a rise
// fails, and a fall fails until the baseline is lowered with
// `UPDATE_UNKNOWN_FACT_BASELINE=1 cargo test -p stella-docx-kernel --test unknown_fact_ratchet`.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;
use std::path::{Path, PathBuf};

use stella_docx_kernel::{
    DocumentPackageProjection, DocumentProjection, DocumentReviewFacts, DocumentStructureFacts,
    DocxLimits, FormattingCompleteness, FormattingFactStatus, InternalParagraphId,
    ParagraphIdentityFacts, ProjectionError, ProjectionOptions, ReviewDetail, ReviewFactLimits,
    ReviewFactSet, ReviewFactUnknownReason, RevisionProjectionStatus, RevisionView,
    StructuralFactSet, project_docx_with_review_facts,
};

const BASELINE: &str = "tests/unknown-fact-baseline.tsv";
const LEGITIMATE: &str = "tests/unknown-fact-legitimate.tsv";
const UPDATE: &str = "UPDATE_UNKNOWN_FACT_BASELINE";
/// Repository directories holding committed packages, relative to the
/// workspace root.
const CORPUS_ROOTS: &[&str] = &[
    "artifacts",
    "packages/core/src/docx",
    "packages/playground/public",
    "parity/fixtures",
    "test/packaged-consumer/public",
    "tests/visual/fixtures",
];

type Key = (String, String, String);
type FixtureKey = (String, String, String, String);

struct Census {
    packages: usize,
    counts: BTreeMap<Key, usize>,
    fixture_counts: BTreeMap<FixtureKey, usize>,
}

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("internal-{}", facts.ordinal))
}

fn workspace_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

fn collect_packages(directory: &Path, packages: &mut Vec<PathBuf>) {
    let entries = std::fs::read_dir(directory)
        .unwrap_or_else(|error| panic!("corpus root {} is readable: {error}", directory.display()));
    for entry in entries {
        let path = entry.unwrap().path();
        let name = path.file_name().unwrap().to_string_lossy();
        if name.starts_with('.') || name == "node_modules" {
            continue;
        }
        let metadata = path.metadata().unwrap_or_else(|error| {
            panic!(
                "corpus entry {} has readable metadata: {error}",
                path.display()
            )
        });
        if metadata.is_dir() {
            collect_packages(&path, packages);
        } else if path
            .extension()
            .is_some_and(|extension| extension == "docx")
        {
            packages.push(path);
        }
    }
}

fn tally(counts: &mut BTreeMap<Key, usize>, view: RevisionView, family: &str, reason: String) {
    *counts
        .entry((format!("{view:?}"), family.to_owned(), reason))
        .or_default() += 1;
}

fn count_unknowns(
    counts: &mut BTreeMap<Key, usize>,
    view: RevisionView,
    projection: &DocumentPackageProjection,
) {
    // Exhaustive fields require a census decision for every new fact family.
    let DocumentPackageProjection {
        document:
            DocumentProjection {
                paragraphs: _paragraphs,
                formatting_completeness,
                revision_status,
                structural_facts:
                    DocumentStructureFacts {
                        indentation,
                        numbering_hierarchy,
                        outline_levels,
                        bookmarks,
                        internal_references,
                    },
            },
        review_facts:
            DocumentReviewFacts {
                revisions: revision_facts,
                comments: comment_facts,
            },
    } = projection;
    // Keep the existing package/reason census unit: a reason shared by several
    // formatting families is counted once. Every family participates explicitly.
    let FormattingCompleteness {
        bold,
        highlight,
        superscript,
        alignment,
    } = formatting_completeness;
    let mut formatting_reasons = BTreeSet::new();
    for status in [bold, highlight, superscript, alignment] {
        match status {
            FormattingFactStatus::Known => {}
            FormattingFactStatus::Unknown(reason) => {
                formatting_reasons.insert(reason);
            }
        }
    }
    for reason in formatting_reasons {
        tally(counts, view, "formatting", format!("{reason:?}"));
    }
    match revision_status {
        RevisionProjectionStatus::Complete => {}
        RevisionProjectionStatus::Incomplete(reasons) => {
            for reason in reasons {
                tally(counts, view, "revision-status", format!("{reason:?}"));
            }
        }
    }
    for (family, reason) in [
        ("indentation", structural_reason(indentation)),
        (
            "numbering-hierarchy",
            structural_reason(numbering_hierarchy),
        ),
        ("outline-levels", structural_reason(outline_levels)),
        ("bookmarks", structural_reason(bookmarks)),
        (
            "internal-references",
            structural_reason(internal_references),
        ),
    ] {
        if let Some(reason) = reason {
            tally(counts, view, family, reason);
        }
    }
    match revision_facts {
        ReviewFactSet::Unknown(reason) => tally(counts, view, "revisions", format!("{reason:?}")),
        ReviewFactSet::Known(revisions) => {
            for revision in revisions {
                if let Some(reason) = review_detail_reason(&revision.content) {
                    tally(
                        counts,
                        view,
                        "revision",
                        format!("{:?}/{reason:?}", revision.kind),
                    );
                }
            }
        }
    }
    match comment_facts {
        ReviewFactSet::Unknown(reason) => tally(counts, view, "comments", format!("{reason:?}")),
        ReviewFactSet::Known(comments) => {
            for comment in comments {
                if let Some(reason) = review_detail_reason(&comment.content) {
                    tally(counts, view, "comment", format!("{reason:?}"));
                }
            }
        }
    }
}

const fn review_detail_reason<T>(detail: &ReviewDetail<T>) -> Option<&ReviewFactUnknownReason> {
    match detail {
        ReviewDetail::Known(_) => None,
        ReviewDetail::Unknown(reason) => Some(reason),
    }
}

fn structural_reason<T>(facts: &StructuralFactSet<T>) -> Option<String> {
    match facts {
        StructuralFactSet::Known(_) => None,
        StructuralFactSet::Unknown(reason) => Some(format!("{reason:?}")),
    }
}

fn census() -> Census {
    let root = workspace_root();
    let mut packages = Vec::new();
    for corpus_root in CORPUS_ROOTS {
        collect_packages(&root.join(corpus_root), &mut packages);
    }
    packages.sort();
    let mut counts = BTreeMap::new();
    let mut fixture_counts = BTreeMap::new();
    for path in &packages {
        let fixture = path
            .strip_prefix(&root)
            .unwrap()
            .to_string_lossy()
            .replace('\\', "/");
        let bytes = std::fs::read(path).unwrap();
        for view in [RevisionView::Current, RevisionView::Original] {
            let projected = project_docx_with_review_facts(
                &bytes,
                DocxLimits::default(),
                ReviewFactLimits::default(),
                ProjectionOptions {
                    revision_view: view,
                    ..ProjectionOptions::default()
                },
                allocate,
            );
            let mut observed = BTreeMap::new();
            match projected {
                Ok(projection) => count_unknowns(&mut observed, view, &projection),
                Err(error) => tally(&mut observed, view, "package", format!("{error:?}")),
            }
            for (key, count) in observed {
                *counts.entry(key.clone()).or_default() += count;
                let (view_label, family, reason) = key;
                fixture_counts.insert((fixture.clone(), view_label, family, reason), count);
            }
        }
    }
    Census {
        packages: packages.len(),
        counts,
        fixture_counts,
    }
}

fn render(packages: usize, counts: &BTreeMap<Key, usize>) -> String {
    let mut output = format!(
        "# Unknown facts across {packages} committed DOCX packages; see tests/unknown_fact_ratchet.rs.\n# view\tfamily\treason\tcount\n"
    );
    for ((view, family, reason), count) in counts {
        writeln!(output, "{view}\t{family}\t{reason}\t{count}").unwrap();
    }
    output
}

fn parse(baseline: &str) -> (usize, BTreeMap<Key, usize>) {
    let packages = baseline
        .lines()
        .next()
        .unwrap()
        .strip_prefix("# Unknown facts across ")
        .unwrap()
        .strip_suffix(" committed DOCX packages; see tests/unknown_fact_ratchet.rs.")
        .unwrap()
        .parse::<usize>()
        .unwrap();
    let counts = baseline
        .lines()
        .filter(|line| !line.starts_with('#') && !line.is_empty())
        .map(|line| {
            let fields = line.split('\t').collect::<Vec<_>>();
            let [view, family, reason, count] = fields.as_slice() else {
                panic!("baseline row has four fields: {line}");
            };
            (
                (
                    (*view).to_owned(),
                    (*family).to_owned(),
                    (*reason).to_owned(),
                ),
                count.parse().unwrap(),
            )
        })
        .collect();
    (packages, counts)
}

fn check_baseline(
    packages: usize,
    counts: &BTreeMap<Key, usize>,
    baseline: &str,
) -> Result<(), String> {
    let (baseline_packages, baseline_counts) = parse(baseline);
    if packages != baseline_packages {
        return Err(format!(
            "corpus package count changed: {baseline_packages} -> {packages}"
        ));
    }
    let mut rose = Vec::new();
    let mut fell = Vec::new();
    let new_keys = counts
        .keys()
        .filter(|key| !baseline_counts.contains_key(*key));
    for key in baseline_counts.keys().chain(new_keys) {
        let before = baseline_counts.get(key).copied().unwrap_or(0);
        let after = counts.get(key).copied().unwrap_or(0);
        if after > before {
            rose.push(format!("{key:?}: {before} -> {after}"));
        } else if after < before {
            fell.push(format!("{key:?}: {before} -> {after}"));
        }
    }
    if !rose.is_empty() {
        return Err(format!("unknown facts rose:\n{}", rose.join("\n")));
    }
    if !fell.is_empty() {
        return Err(format!(
            "unknown facts fell; lower the baseline with {UPDATE}=1:\n{}",
            fell.join("\n")
        ));
    }
    Ok(())
}

proptest::proptest! {
    /// Corpus coverage is checked even when no package has an unknown fact.
    #[test]
    fn corpus_coverage_requires_the_recorded_package_count(packages in 2_usize..128) {
        let counts = BTreeMap::new();
        let baseline = render(packages, &counts);
        proptest::prop_assert!(check_baseline(packages, &counts, &baseline).is_ok());
        proptest::prop_assert!(check_baseline(packages - 1, &counts, &baseline).is_err());
        proptest::prop_assert!(check_baseline(packages + 1, &counts, &baseline).is_err());
    }
}

#[cfg(unix)]
#[test]
fn corpus_traversal_reports_unreadable_entry_metadata() {
    use std::os::unix::fs::symlink;

    let directory = std::env::temp_dir().join(format!("folio-census-{}", std::process::id()));
    std::fs::create_dir_all(&directory).unwrap();
    symlink(directory.join("absent"), directory.join("entry")).unwrap();
    let result = std::panic::catch_unwind(|| collect_packages(&directory, &mut Vec::new()));
    std::fs::remove_dir_all(directory).unwrap();
    assert!(
        result.is_err(),
        "every corpus entry must have readable metadata"
    );
}

fn parse_legitimate_unknowns(contents: &str) -> BTreeMap<FixtureKey, usize> {
    let mut permitted = BTreeMap::new();
    for line in contents
        .lines()
        .filter(|line| !line.starts_with('#') && !line.is_empty())
    {
        let fields = line.split('\t').collect::<Vec<_>>();
        let [fixture, view, family, reason, count, rationale] = fields.as_slice() else {
            panic!("legitimate unknown row has six fields: {line}");
        };
        assert!(
            matches!(*view, "Current" | "Original"),
            "invalid revision view: {line}"
        );
        assert!(
            !fixture.is_empty()
                && !family.is_empty()
                && !reason.is_empty()
                && !rationale.trim().is_empty(),
            "legitimate unknown needs a fixture, family, reason, and rationale: {line}"
        );
        let count = count
            .parse::<usize>()
            .expect("legitimate count is an integer");
        assert!(
            count > 0,
            "legitimate unknown count must be positive: {line}"
        );
        let key = (
            (*fixture).to_owned(),
            (*view).to_owned(),
            (*family).to_owned(),
            (*reason).to_owned(),
        );
        assert!(
            permitted.insert(key, count).is_none(),
            "duplicate legitimate unknown tuple: {line}"
        );
    }
    permitted
}

fn legitimate_unknown_violations(
    actual: &BTreeMap<FixtureKey, usize>,
    permitted: &BTreeMap<FixtureKey, usize>,
) -> Vec<String> {
    let mut violations = Vec::new();
    for (key, count) in actual {
        match permitted.get(key) {
            None => violations.push(format!("unlisted unknown {key:?}: {count}")),
            Some(expected) if count != expected => violations.push(format!(
                "audited count changed {key:?}: {expected} -> {count}"
            )),
            Some(_) => {}
        }
    }
    for (key, count) in permitted {
        if !actual.contains_key(key) {
            violations.push(format!(
                "audited unknown is absent {key:?}: expected {count}"
            ));
        }
    }
    violations
}

#[test]
fn legitimate_detector_accepts_exact_audited_counts() {
    let allowed = parse_legitimate_unknowns(
        "fixture-a.docx\tCurrent\tcomment\tUnsupportedLocation\t4\tComments have no document anchors.\n",
    );
    assert!(legitimate_unknown_violations(&allowed, &allowed).is_empty());
}

#[test]
fn legitimate_detector_rejects_replaced_fixtures_without_a_total_increase() {
    let allowed = parse_legitimate_unknowns(
        "fixture-a.docx\tCurrent\tcomment\tUnsupportedLocation\t4\tComments have no document anchors.\n",
    );
    let replaced = parse_legitimate_unknowns(
        "fixture-b.docx\tCurrent\tcomment\tUnsupportedLocation\t4\tSynthetic replacement.\n",
    );
    assert_eq!(
        allowed.values().sum::<usize>(),
        replaced.values().sum::<usize>()
    );
    let violations = legitimate_unknown_violations(&replaced, &allowed);
    assert!(
        violations
            .iter()
            .any(|violation| violation.starts_with("unlisted unknown"))
    );
}

#[test]
fn legitimate_detector_rejects_unlisted_views_families_and_reasons() {
    let allowed = parse_legitimate_unknowns(
        "fixture-a.docx\tCurrent\tcomment\tUnsupportedLocation\t4\tComments have no document anchors.\n",
    );
    for row in [
        "fixture-a.docx\tOriginal\tcomment\tUnsupportedLocation\t4\tSynthetic replacement.",
        "fixture-a.docx\tCurrent\trevision\tUnsupportedLocation\t4\tSynthetic replacement.",
        "fixture-a.docx\tCurrent\tcomment\tPayloadBudgetExceeded\t4\tSynthetic replacement.",
    ] {
        let observed = parse_legitimate_unknowns(row);
        let violations = legitimate_unknown_violations(&observed, &allowed);
        assert!(
            violations
                .iter()
                .any(|violation| violation.starts_with("unlisted unknown")),
            "{row}"
        );
    }
}

#[test]
fn legitimate_detector_requires_counts_and_list_to_stay_current() {
    let allowed = parse_legitimate_unknowns(
        "fixture-a.docx\tCurrent\tcomment\tUnsupportedLocation\t4\tComments have no document anchors.\n",
    );
    for count in [3, 5] {
        let observed = parse_legitimate_unknowns(&format!(
            "fixture-a.docx\tCurrent\tcomment\tUnsupportedLocation\t{count}\tSynthetic count change."
        ));
        assert!(!legitimate_unknown_violations(&observed, &allowed).is_empty());
    }
    assert!(!legitimate_unknown_violations(&BTreeMap::new(), &allowed).is_empty());
}

#[test]
fn unknown_fact_counts_only_fall() {
    let Census {
        packages,
        counts,
        fixture_counts,
    } = census();
    assert!(packages > 0, "the corpus roots hold packages");
    let legitimate_path = Path::new(env!("CARGO_MANIFEST_DIR")).join(LEGITIMATE);
    let permitted = parse_legitimate_unknowns(&std::fs::read_to_string(legitimate_path).unwrap());
    let violations = legitimate_unknown_violations(&fixture_counts, &permitted);
    assert!(
        violations.is_empty(),
        "unknown facts differ from the audited list:\n{}",
        violations.join("\n")
    );
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join(BASELINE);
    if std::env::var_os(UPDATE).is_some() {
        std::fs::write(&path, render(packages, &counts)).unwrap();
        return;
    }
    check_baseline(packages, &counts, &std::fs::read_to_string(&path).unwrap()).unwrap();
}
