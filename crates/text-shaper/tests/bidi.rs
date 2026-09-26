//! Conformance of the bidirectional algorithm against the Unicode test data.
//!
//! The fixture is a subset of `BidiCharacterTest.txt`: every case through the
//! miscellaneous sections (which hold all of the isolate, override and bracket
//! cases) and a sample of the permutation sections.

#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used
)]
// Assertion-style failure is the point of a test, and indexing happens only
// after the shape of the collection under test is established.

use stella_text_shaper::{BaseDirection, BidiLine, resolve_bidi};

const FIXTURE: &str = include_str!("fixtures/BidiCharacterTest-subset.txt");

/// Levels as the test data writes them, with `x` for characters rule X9
/// removes, and the visual order with those characters skipped.
fn as_test_data(line: &BidiLine, removed: &[bool]) -> (String, String) {
    let levels = line
        .levels
        .iter()
        .zip(removed)
        .map(|(level, removed)| {
            if *removed {
                "x".to_owned()
            } else {
                level.to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(" ");
    let order = line
        .visual_order
        .iter()
        .filter(|index| !removed[**index])
        .map(ToString::to_string)
        .collect::<Vec<_>>()
        .join(" ");
    (levels, order)
}

#[test]
fn every_case_of_the_fixture_resolves_as_the_standard_does() {
    let mut cases = 0_usize;
    let mut failures = Vec::new();
    for (number, case) in FIXTURE.lines().enumerate() {
        if case.is_empty() || case.starts_with('#') {
            continue;
        }
        let fields: Vec<&str> = case.split(';').collect();
        assert_eq!(fields.len(), 5, "line {} has five fields", number + 1);
        let text: String = fields[0]
            .split(' ')
            .map(|hex| {
                char::from_u32(u32::from_str_radix(hex, 16).expect("a hex code point"))
                    .expect("a scalar value")
            })
            .collect();
        let base = match fields[1] {
            "0" => BaseDirection::LeftToRight,
            "1" => BaseDirection::RightToLeft,
            "2" => BaseDirection::Auto,
            other => panic!("unknown paragraph direction {other}"),
        };
        let removed: Vec<bool> = fields[3].split(' ').map(|level| level == "x").collect();

        let line = resolve_bidi(&text, base);
        let (levels, order) = as_test_data(&line, &removed);
        cases += 1;
        if line.paragraph_level.to_string() != fields[2]
            || levels != fields[3]
            || order != fields[4]
        {
            failures.push(format!(
                "line {}: {case}\n  got {};{levels};{order}",
                number + 1,
                line.paragraph_level
            ));
        }
    }
    assert!(cases > 500, "the fixture holds the sampled cases: {cases}");
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn an_isolated_right_to_left_name_keeps_a_following_number_outside_it() {
    // A left-to-right sentence with a right-to-left name inserted and a date
    // after it. Without the isolate, the digits join the name's run and read
    // before it; inside one, the name is a neutral to its surroundings.
    let visual = |text: &str| -> String {
        let line = resolve_bidi(text, BaseDirection::LeftToRight);
        let characters: Vec<char> = text.chars().collect();
        line.visual_order
            .iter()
            .map(|index| characters[*index])
            .filter(|character| !('\u{2066}'..='\u{2069}').contains(character))
            .collect()
    };

    assert_eq!(
        visual("signed by \u{645}\u{62d}\u{645}\u{62f} 2026"),
        "signed by 2026 \u{62f}\u{645}\u{62d}\u{645}",
        "bare, the date joins the name's run"
    );
    assert_eq!(
        visual("signed by \u{2068}\u{645}\u{62d}\u{645}\u{62f}\u{2069} 2026"),
        "signed by \u{62f}\u{645}\u{62d}\u{645} 2026",
        "isolated, the name reverses in place and the date stays last"
    );
}

#[test]
fn an_automatic_direction_follows_the_first_strong_character() {
    assert_eq!(
        resolve_bidi("\u{5d0} abc", BaseDirection::Auto).paragraph_level,
        1,
        "Hebrew first"
    );
    assert_eq!(
        resolve_bidi("123 abc \u{5d0}", BaseDirection::Auto).paragraph_level,
        0,
        "Latin first"
    );
    // Rule P2 skips isolates when it looks for the first strong character.
    assert_eq!(
        resolve_bidi("\u{2067}abc\u{2069} \u{5d0}", BaseDirection::Auto).paragraph_level,
        1,
        "an isolate does not decide the paragraph"
    );
}

#[test]
fn empty_text_resolves_to_nothing() {
    let line = resolve_bidi("", BaseDirection::RightToLeft);
    assert_eq!(line.paragraph_level, 1, "the requested level");
    assert!(line.levels.is_empty(), "no levels");
    assert!(line.visual_order.is_empty(), "no order");
}
