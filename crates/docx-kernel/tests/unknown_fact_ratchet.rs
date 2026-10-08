#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::panic,
    clippy::unwrap_used
)]
// Unknown-fact ratchet over the repository's committed DOCX packages.
//
// Count typed unknown facts per (view, family, reason) across every committed
// package and compare with `unknown-fact-baseline.tsv`. A count may only fall: a rise
// fails, and a fall fails until the baseline is lowered with
// `UPDATE_UNKNOWN_FACT_BASELINE=1 cargo test -p stella-docx-kernel --test unknown_fact_ratchet`.

use std::collections::BTreeMap;
use std::fmt::Write as _;
use std::path::{Path, PathBuf};

use stella_docx_kernel::{
    DocumentPackageProjection, DocxLimits, FormattingProjectionStatus, InternalParagraphId,
    ParagraphIdentityFacts, ProjectionError, ProjectionOptions, ReviewDetail, ReviewFactLimits,
    ReviewFactSet, RevisionProjectionStatus, RevisionView, StructuralFactSet,
    project_docx_with_review_facts,
};

const BASELINE: &str = "tests/unknown-fact-baseline.tsv";
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
    let document = &projection.document;
    if let FormattingProjectionStatus::Incomplete(reason) = document.formatting_status {
        tally(counts, view, "formatting", format!("{reason:?}"));
    }
    if let RevisionProjectionStatus::Incomplete(reasons) = &document.revision_status {
        for reason in reasons {
            tally(counts, view, "revision-status", format!("{reason:?}"));
        }
    }
    let facts = &document.structural_facts;
    for (family, reason) in [
        ("indentation", structural_reason(&facts.indentation)),
        (
            "numbering-hierarchy",
            structural_reason(&facts.numbering_hierarchy),
        ),
        ("outline-levels", structural_reason(&facts.outline_levels)),
        ("bookmarks", structural_reason(&facts.bookmarks)),
        (
            "internal-references",
            structural_reason(&facts.internal_references),
        ),
    ] {
        if let Some(reason) = reason {
            tally(counts, view, family, reason);
        }
    }
    match &projection.review_facts.revisions {
        ReviewFactSet::Unknown(reason) => tally(counts, view, "revisions", format!("{reason:?}")),
        ReviewFactSet::Known(revisions) => {
            for revision in revisions {
                if let ReviewDetail::Unknown(reason) = revision.content {
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
    match &projection.review_facts.comments {
        ReviewFactSet::Unknown(reason) => tally(counts, view, "comments", format!("{reason:?}")),
        ReviewFactSet::Known(comments) => {
            for comment in comments {
                if let ReviewDetail::Unknown(reason) = comment.content {
                    tally(counts, view, "comment", format!("{reason:?}"));
                }
            }
        }
    }
}

fn structural_reason<T>(facts: &StructuralFactSet<T>) -> Option<String> {
    match facts {
        StructuralFactSet::Known(_) => None,
        StructuralFactSet::Unknown(reason) => Some(format!("{reason:?}")),
    }
}

fn census() -> (usize, BTreeMap<Key, usize>) {
    let root = workspace_root();
    let mut packages = Vec::new();
    for corpus_root in CORPUS_ROOTS {
        collect_packages(&root.join(corpus_root), &mut packages);
    }
    packages.sort();
    let mut counts = BTreeMap::new();
    for path in &packages {
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
            match projected {
                Ok(projection) => count_unknowns(&mut counts, view, &projection),
                Err(error) => tally(&mut counts, view, "package", format!("{error:?}")),
            }
        }
    }
    (packages.len(), counts)
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

#[test]
fn unknown_fact_counts_only_fall() {
    let (packages, counts) = census();
    assert!(packages > 0, "the corpus roots hold packages");
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join(BASELINE);
    if std::env::var_os(UPDATE).is_some() {
        std::fs::write(&path, render(packages, &counts)).unwrap();
        return;
    }
    check_baseline(packages, &counts, &std::fs::read_to_string(&path).unwrap()).unwrap();
}
