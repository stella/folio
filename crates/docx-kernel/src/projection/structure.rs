use std::collections::{HashMap, HashSet};

use crate::ProjectionError;
use crate::projection::numbering::NumberingCatalog;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum StructuralFactUnknownReason {
    DocumentPartOnly,
    StylesPartUnavailable,
    UnsupportedStyles,
    UnsupportedNumbering,
    IncompleteBookmarkRanges,
    UnsupportedInternalReferences,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum StructuralFactSet<T> {
    Known(Vec<T>),
    Unknown(StructuralFactUnknownReason),
}

impl<T> StructuralFactSet<T> {
    pub fn known(items: impl IntoIterator<Item = T>) -> Self {
        Self::Known(items.into_iter().collect())
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct ParagraphIndentation {
    pub first_line_twips: Option<i32>,
    pub hanging_twips: Option<i32>,
    pub left_twips: Option<i32>,
    pub right_twips: Option<i32>,
    pub start_twips: Option<i32>,
    pub end_twips: Option<i32>,
    pub first_line_chars_hundredths: Option<i32>,
    pub hanging_chars_hundredths: Option<i32>,
    pub left_chars_hundredths: Option<i32>,
    pub right_chars_hundredths: Option<i32>,
    pub start_chars_hundredths: Option<i32>,
    pub end_chars_hundredths: Option<i32>,
}

impl ParagraphIndentation {
    fn is_empty(self) -> bool {
        self == Self::default()
    }

    fn inherit(self, child: Self) -> Self {
        let child_has_hanging =
            child.hanging_twips.is_some() || child.hanging_chars_hundredths.is_some();
        let child_has_first_line = !child_has_hanging
            && (child.first_line_twips.is_some() || child.first_line_chars_hundredths.is_some());
        let parent_first_line = if child_has_hanging {
            (None, None)
        } else {
            (self.first_line_twips, self.first_line_chars_hundredths)
        };
        let parent_hanging = if child_has_first_line {
            (None, None)
        } else {
            (self.hanging_twips, self.hanging_chars_hundredths)
        };
        let child_first_line = if child_has_hanging {
            (None, None)
        } else {
            (child.first_line_twips, child.first_line_chars_hundredths)
        };
        let (first_line_twips, first_line_chars_hundredths) = inherit_character_indent(
            parent_first_line.0,
            parent_first_line.1,
            child_first_line.0,
            child_first_line.1,
        );
        let (hanging_twips, hanging_chars_hundredths) = inherit_character_indent(
            parent_hanging.0,
            parent_hanging.1,
            child.hanging_twips,
            child.hanging_chars_hundredths,
        );
        let (mut left_twips, left_chars_hundredths) = inherit_character_indent(
            self.left_twips,
            self.left_chars_hundredths,
            child.left_twips,
            child.left_chars_hundredths,
        );
        let (mut right_twips, right_chars_hundredths) = inherit_character_indent(
            self.right_twips,
            self.right_chars_hundredths,
            child.right_twips,
            child.right_chars_hundredths,
        );
        let (mut start_twips, start_chars_hundredths) = inherit_character_indent(
            self.start_twips,
            self.start_chars_hundredths,
            child.start_twips,
            child.start_chars_hundredths,
        );
        let (mut end_twips, end_chars_hundredths) = inherit_character_indent(
            self.end_twips,
            self.end_chars_hundredths,
            child.end_twips,
            child.end_chars_hundredths,
        );
        if left_chars_hundredths.is_some() || start_chars_hundredths.is_some() {
            left_twips = None;
            start_twips = None;
        }
        if right_chars_hundredths.is_some() || end_chars_hundredths.is_some() {
            right_twips = None;
            end_twips = None;
        }
        Self {
            first_line_twips,
            hanging_twips,
            left_twips,
            right_twips,
            start_twips,
            end_twips,
            first_line_chars_hundredths,
            hanging_chars_hundredths,
            left_chars_hundredths,
            right_chars_hundredths,
            start_chars_hundredths,
            end_chars_hundredths,
        }
    }

    fn inherit_numbered_direct(self, mut direct: Self) -> Self {
        if direct.first_line_twips == Some(0) {
            direct.first_line_twips = None;
        }
        if direct.hanging_twips == Some(0) {
            direct.hanging_twips = None;
        }
        self.inherit(direct)
    }
}

fn inherit_character_indent(
    parent_twips: Option<i32>,
    parent_chars: Option<i32>,
    child_twips: Option<i32>,
    child_chars: Option<i32>,
) -> (Option<i32>, Option<i32>) {
    match child_chars {
        Some(0) => (child_twips.or(parent_twips), None),
        Some(value) => (None, Some(value)),
        None => parent_chars.filter(|value| *value != 0).map_or_else(
            || (child_twips.or(parent_twips), None),
            |value| (None, Some(value)),
        ),
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ParagraphIndentationFact {
    pub paragraph_ordinal: usize,
    pub value: ParagraphIndentation,
}

/// Projected `w:jc/@w:val` tokens.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ParagraphAlignmentValue {
    Center,
    Justify,
    Left,
    Right,
}

/// Cascade level that supplied a paragraph's effective `w:jc`.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ParagraphAlignmentSource {
    /// `w:pPr/w:jc` on the paragraph itself.
    Direct,
    /// The `w:pStyle` chain, or `w:docDefaults/w:pPrDefault/w:pPr/w:jc`.
    Style,
}

/// Effective paragraph alignment resolved for one paragraph.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ParagraphAlignmentFact {
    pub value: ParagraphAlignmentValue,
    pub source: ParagraphAlignmentSource,
}

/// A `w:jc` element observed at one cascade level.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum ParagraphAlignmentSetting {
    Supported(ParagraphAlignmentValue),
    /// A `w:val` token outside the projected set, or a `w:jc` without a usable
    /// `w:val`. The element is present, so it still shadows lower cascade
    /// levels; it projects no value, because a lower level's value would be a
    /// guess at what the token renders as.
    Unsupported,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ParagraphOutlineLevelFact {
    pub paragraph_ordinal: usize,
    pub outline_level: u8,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct NumberingHierarchyFact {
    pub paragraph_ordinal: usize,
    pub parent_paragraph_ordinal: Option<usize>,
    pub child_paragraph_ordinals: Vec<usize>,
}

#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd)]
pub enum SpanCoverage {
    Complete,
    ContinuesBefore,
    ContinuesAfter,
    ContinuesBeforeAndAfter,
}

#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd)]
pub struct StructuralSpan {
    pub start_utf8: u32,
    pub end_utf8: u32,
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub coverage: SpanCoverage,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BookmarkFact {
    pub paragraph_ordinal: usize,
    pub bookmark_id: u32,
    pub name: String,
    pub span: StructuralSpan,
}

#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd)]
pub enum InternalReferenceRole {
    Source,
    Target,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct InternalReferenceFact {
    pub paragraph_ordinal: usize,
    pub reference_id: String,
    pub role: InternalReferenceRole,
    pub span: StructuralSpan,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DocumentStructureFacts {
    pub indentation: StructuralFactSet<ParagraphIndentationFact>,
    pub numbering_hierarchy: StructuralFactSet<NumberingHierarchyFact>,
    pub bookmarks: StructuralFactSet<BookmarkFact>,
    pub internal_references: StructuralFactSet<InternalReferenceFact>,
    pub outline_levels: StructuralFactSet<ParagraphOutlineLevelFact>,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub(super) struct NumberingProperties {
    pub num_id: Option<u32>,
    pub level: Option<u8>,
    pub present: bool,
}

impl NumberingProperties {
    fn inherit(self, child: Self) -> Self {
        if !child.present {
            return self;
        }
        Self {
            num_id: child.num_id.or(self.num_id),
            level: child.level.or(self.level),
            present: true,
        }
    }
}

#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub(super) struct ParagraphProperties {
    pub style_id: Option<String>,
    pub outline_level: Option<u8>,
    pub indentation: ParagraphIndentation,
    pub numbering: NumberingProperties,
    pub alignment: Option<ParagraphAlignmentSetting>,
}

#[derive(Clone, Debug)]
pub(super) struct RawBlockPoint {
    pub paragraph: usize,
    pub utf8: u32,
    pub utf16: u32,
}

#[derive(Clone, Debug)]
pub(super) struct RawBookmarkRange {
    pub id: u32,
    pub name: String,
    pub start: RawBlockPoint,
    pub end: RawBlockPoint,
}

#[derive(Clone, Debug)]
pub(super) struct RawInternalReference {
    pub reference_id: String,
    pub source: RawBlockPoint,
}

#[derive(Clone, Debug)]
pub(super) struct StyleDefinition {
    pub based_on: Option<String>,
    pub properties: ParagraphProperties,
    pub text_properties: TextProperties,
    pub kind: StyleKind,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum StyleKind {
    Character,
    Paragraph,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub(super) struct TextProperties {
    pub bold: Option<bool>,
    pub complex_script_bold: Option<bool>,
    pub force_complex_script: Option<bool>,
    pub right_to_left: Option<bool>,
    pub highlighted: Option<bool>,
    pub superscript: Option<bool>,
}

impl TextProperties {
    const fn inherit(self, child: Self) -> Self {
        Self {
            bold: if child.bold.is_some() {
                child.bold
            } else {
                self.bold
            },
            complex_script_bold: if child.complex_script_bold.is_some() {
                child.complex_script_bold
            } else {
                self.complex_script_bold
            },
            force_complex_script: if child.force_complex_script.is_some() {
                child.force_complex_script
            } else {
                self.force_complex_script
            },
            right_to_left: if child.right_to_left.is_some() {
                child.right_to_left
            } else {
                self.right_to_left
            },
            highlighted: if child.highlighted.is_some() {
                child.highlighted
            } else {
                self.highlighted
            },
            superscript: if child.superscript.is_some() {
                child.superscript
            } else {
                self.superscript
            },
        }
    }

    fn inherit_style(self, child: Self) -> Self {
        let mut inherited = self.inherit(child);
        inherited.bold = inherit_style_toggle(self.bold, child.bold);
        inherited.complex_script_bold =
            inherit_style_toggle(self.complex_script_bold, child.complex_script_bold);
        inherited
    }
}

fn inherit_style_toggle(inherited: Option<bool>, child: Option<bool>) -> Option<bool> {
    match child {
        Some(true) => Some(!inherited.unwrap_or(false)),
        Some(false) | None => inherited,
    }
}

#[derive(Clone, Debug, Default)]
pub(super) struct StyleSheet {
    pub default_style_id: Option<String>,
    pub document_defaults: ParagraphProperties,
    pub document_text_defaults: TextProperties,
    pub styles: HashMap<String, StyleDefinition>,
    resolved_character_text: HashMap<String, TextProperties>,
    resolved_paragraph_styles: HashMap<String, ResolvedParagraphStyle>,
}

#[derive(Clone, Copy, Debug, Default)]
struct ResolvedParagraphStyle {
    text: TextProperties,
    properties: ResolvedParagraphStyleProperties,
}

impl ResolvedParagraphStyle {
    fn inherit(self, child: &StyleDefinition) -> Self {
        Self {
            text: self.text.inherit_style(child.text_properties),
            properties: self.properties.inherit(&child.properties),
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct ResolvedParagraphStyleProperties {
    indentation: ParagraphIndentation,
    numbering: NumberingProperties,
    outline_level: Option<u8>,
    alignment: Option<ParagraphAlignmentSetting>,
}

impl ResolvedParagraphStyleProperties {
    fn inherit(self, child: &ParagraphProperties) -> Self {
        Self {
            indentation: self.indentation.inherit(child.indentation),
            numbering: self.numbering.inherit(child.numbering),
            outline_level: child.outline_level.or(self.outline_level),
            alignment: child.alignment.or(self.alignment),
        }
    }
}

#[derive(Clone, Debug)]
struct ResolvedParagraphProperties {
    indentation: Result<ParagraphIndentation, StructuralFactUnknownReason>,
    numbering: NumberingProperties,
    outline_level: Option<u8>,
}

impl StyleSheet {
    pub(super) fn prepare_styles(&mut self) {
        (self.resolved_character_text, _) = self.resolve_character_text_styles();
        (self.resolved_paragraph_styles, _) = self.resolve_paragraph_styles();
    }

    pub(super) fn resolve_text(
        &self,
        paragraph_style_id: Option<&str>,
        character_style_id: Option<&str>,
        direct: TextProperties,
    ) -> Result<TextProperties, ()> {
        let paragraph_style_id = paragraph_style_id.or(self.default_style_id.as_deref());
        let paragraph = paragraph_style_id
            .map(|style_id| {
                self.resolved_paragraph_styles
                    .get(style_id)
                    .map(|style| style.text)
                    .ok_or(())
            })
            .transpose()?
            .unwrap_or_default();
        let character = character_style_id
            .map(|style_id| {
                self.resolved_character_text
                    .get(style_id)
                    .copied()
                    .ok_or(())
            })
            .transpose()?
            .unwrap_or_default();
        let mut effective = self
            .document_text_defaults
            .inherit(paragraph)
            .inherit(character)
            .inherit(direct);
        effective.bold = Some(resolve_word_toggle(
            self.document_text_defaults.bold,
            paragraph.bold,
            character.bold,
            direct.bold,
        ));
        effective.complex_script_bold = Some(resolve_word_toggle(
            self.document_text_defaults.complex_script_bold,
            paragraph.complex_script_bold,
            character.complex_script_bold,
            direct.complex_script_bold,
        ));
        Ok(effective)
    }

    fn resolve_character_text_styles(&self) -> (HashMap<String, TextProperties>, usize) {
        let mut resolved = HashMap::new();
        let mut unresolved = HashSet::new();
        let mut resolution_steps = 0usize;
        for root_id in self.styles.iter().filter_map(|(style_id, style)| {
            (style.kind == StyleKind::Character).then_some(style_id)
        }) {
            if resolved.contains_key(root_id) || unresolved.contains(root_id) {
                continue;
            }
            let mut chain = Vec::new();
            let mut seen = HashSet::new();
            let mut current = Some(root_id.as_str());
            let mut inherited = TextProperties::default();
            let mut valid = true;
            while let Some(style_id) = current {
                resolution_steps = resolution_steps.saturating_add(1);
                if let Some(properties) = resolved.get(style_id).copied() {
                    inherited = properties;
                    break;
                }
                if unresolved.contains(style_id) {
                    valid = false;
                    break;
                }
                if !seen.insert(style_id) {
                    valid = false;
                    break;
                }
                let Some(style) = self.styles.get(style_id) else {
                    // Word ignores a missing basedOn target. A missing initially
                    // selected style is still unsupported.
                    valid = !chain.is_empty();
                    break;
                };
                if style.kind != StyleKind::Character {
                    // basedOn inheritance cannot cross style kinds; Word treats
                    // the current style as a root instead of discarding it.
                    valid = !chain.is_empty();
                    break;
                }
                chain.push(style_id);
                current = style.based_on.as_deref();
            }
            if !valid {
                unresolved.extend(chain.into_iter().map(str::to_owned));
                continue;
            }
            for style_id in chain.into_iter().rev() {
                // Every ID in the chain was validated above.
                if let Some(style) = self.styles.get(style_id) {
                    inherited = inherited.inherit_style(style.text_properties);
                    resolved.insert(style_id.to_owned(), inherited);
                }
            }
        }
        debug_assert!(
            resolution_steps <= self.styles.len().saturating_mul(2).saturating_add(1),
            "memoized style resolution exceeded its linear work budget"
        );
        (resolved, resolution_steps)
    }

    fn resolve_paragraph_styles(&self) -> (HashMap<String, ResolvedParagraphStyle>, usize) {
        let mut resolved = HashMap::new();
        let mut unresolved = HashSet::new();
        let mut resolution_steps = 0usize;
        for root_id in self.styles.iter().filter_map(|(style_id, style)| {
            (style.kind == StyleKind::Paragraph).then_some(style_id)
        }) {
            if resolved.contains_key(root_id) || unresolved.contains(root_id) {
                continue;
            }
            let mut chain = Vec::new();
            let mut seen = HashSet::new();
            let mut current = Some(root_id.as_str());
            let mut inherited = ResolvedParagraphStyle::default();
            let mut valid = true;
            while let Some(style_id) = current {
                resolution_steps = resolution_steps.saturating_add(1);
                if let Some(properties) = resolved.get(style_id).copied() {
                    inherited = properties;
                    break;
                }
                if unresolved.contains(style_id) || !seen.insert(style_id) {
                    valid = false;
                    break;
                }
                let Some(style) = self.styles.get(style_id) else {
                    valid = !chain.is_empty();
                    break;
                };
                if style.kind != StyleKind::Paragraph {
                    valid = !chain.is_empty();
                    break;
                }
                chain.push(style_id);
                current = style.based_on.as_deref();
            }
            if !valid {
                unresolved.extend(chain.into_iter().map(str::to_owned));
                continue;
            }
            for style_id in chain.into_iter().rev() {
                if let Some(style) = self.styles.get(style_id) {
                    inherited = inherited.inherit(style);
                    resolved.insert(style_id.to_owned(), inherited);
                }
            }
        }
        debug_assert!(
            resolution_steps <= self.styles.len().saturating_mul(2).saturating_add(1),
            "memoized paragraph style resolution exceeded its linear work budget"
        );
        (resolved, resolution_steps)
    }

    /// Resolves `w:jc` from the paragraph style chain, then `w:docDefaults`.
    ///
    /// An initially selected style that does not resolve (a missing `w:pStyle`
    /// target, a `w:basedOn` cycle, or a kind-crossing chain) yields no
    /// style-tier alignment at all, matching [`Self::resolve`] and
    /// [`Self::resolve_text`]. Word falls back to Normal there, so reporting
    /// `w:docDefaults` instead would claim a value the document never resolves
    /// to.
    fn style_alignment(&self, direct: &ParagraphProperties) -> Option<ParagraphAlignmentSetting> {
        let initial_style_id = direct.style_id.as_ref().or(self.default_style_id.as_ref());
        let style = match initial_style_id {
            Some(style_id) => Some(self.resolved_paragraph_styles.get(style_id)?),
            None => None,
        };
        style
            .and_then(|style| style.properties.alignment)
            .or(self.document_defaults.alignment)
    }

    pub(super) fn paragraph_uses_numbering(
        &self,
        direct: &ParagraphProperties,
    ) -> Result<bool, StructuralFactUnknownReason> {
        self.resolve(
            direct,
            Err(StructuralFactUnknownReason::UnsupportedNumbering),
        )
        .map(|resolved| resolved.numbering.present && resolved.numbering.num_id != Some(0))
    }

    fn resolve(
        &self,
        direct: &ParagraphProperties,
        numbering_catalog: Result<&NumberingCatalog, StructuralFactUnknownReason>,
    ) -> Result<ResolvedParagraphProperties, StructuralFactUnknownReason> {
        let initial_style_id = direct.style_id.as_ref().or(self.default_style_id.as_ref());
        let style = initial_style_id
            .map(|style_id| {
                self.resolved_paragraph_styles
                    .get(style_id)
                    .map(|style| style.properties)
                    .ok_or(StructuralFactUnknownReason::UnsupportedStyles)
            })
            .transpose()?
            .unwrap_or_default();
        let mut numbering = self.document_defaults.numbering.inherit(style.numbering);
        let mut outline_level = style.outline_level.or(self.document_defaults.outline_level);
        let inherited_numbering = numbering;
        numbering = numbering.inherit(direct.numbering);
        outline_level = direct.outline_level.or(outline_level);

        let numbered = numbering.present && numbering.num_id != Some(0);
        let numbering_is_direct = direct.numbering.present;
        let numbering_removed = numbering_is_direct
            && direct.numbering.num_id == Some(0)
            && inherited_numbering.present
            && inherited_numbering.num_id != Some(0);
        let indentation = if numbered {
            numbering
                .num_id
                .ok_or(StructuralFactUnknownReason::UnsupportedNumbering)
                .and_then(|num_id| {
                    numbering_catalog.and_then(|catalog| {
                        catalog.indentation(num_id, numbering.level.unwrap_or(0))
                    })
                })
        } else {
            Ok(ParagraphIndentation::default())
        }
        .map(|level_indentation| {
            let mut indentation = if numbering_removed {
                ParagraphIndentation::default()
            } else {
                self.document_defaults.indentation
            };
            if !numbering_is_direct && !numbering_removed {
                indentation = indentation.inherit(level_indentation);
            }
            if !numbering_removed {
                indentation = indentation.inherit(style.indentation);
            }
            if numbering_is_direct {
                indentation = indentation.inherit(level_indentation);
            }
            if numbered {
                indentation.inherit_numbered_direct(direct.indentation)
            } else {
                indentation.inherit(direct.indentation)
            }
        });
        Ok(ResolvedParagraphProperties {
            indentation,
            numbering,
            outline_level,
        })
    }
}

/// Resolves the effective paragraph alignment for one paragraph.
///
/// Direct `w:pPr/w:jc` wins; otherwise the paragraph style chain and
/// `w:docDefaults` apply. Without a style sheet only direct alignment is
/// projected, which never claims a style-dependent value.
///
/// Two cascade tiers are out of scope: a table style's `w:tblStyle` ->
/// `w:pPr/w:jc` and a numbering level's `w:lvl/w:pPr/w:jc`. Neither is
/// consulted, so a paragraph whose effective alignment comes from either tier
/// projects the lower-tier `w:docDefaults` or style value instead.
pub(super) fn resolve_paragraph_alignment(
    styles: Result<&StyleSheet, StructuralFactUnknownReason>,
    direct: &ParagraphProperties,
) -> Option<ParagraphAlignmentFact> {
    if let Some(setting) = direct.alignment {
        return alignment_fact(setting, ParagraphAlignmentSource::Direct);
    }
    styles
        .ok()?
        .style_alignment(direct)
        .and_then(|setting| alignment_fact(setting, ParagraphAlignmentSource::Style))
}

const fn alignment_fact(
    setting: ParagraphAlignmentSetting,
    source: ParagraphAlignmentSource,
) -> Option<ParagraphAlignmentFact> {
    match setting {
        ParagraphAlignmentSetting::Supported(value) => {
            Some(ParagraphAlignmentFact { value, source })
        }
        ParagraphAlignmentSetting::Unsupported => None,
    }
}

fn resolve_word_toggle(
    document_default: Option<bool>,
    paragraph: Option<bool>,
    character: Option<bool>,
    direct: Option<bool>,
) -> bool {
    if let Some(value) = direct {
        return value;
    }
    document_default.unwrap_or(false) ^ paragraph.unwrap_or(false) ^ character.unwrap_or(false)
}

#[derive(Clone, Copy)]
pub(super) struct RawStructureInput<'a> {
    pub paragraph_texts: &'a [&'a str],
    pub properties: &'a [&'a ParagraphProperties],
    pub bookmarks: Result<&'a [RawBookmarkRange], StructuralFactUnknownReason>,
    pub references: Result<&'a [RawInternalReference], StructuralFactUnknownReason>,
}

struct StructuralFactBudget {
    remaining: usize,
}

impl StructuralFactBudget {
    const fn new(maximum_facts: usize) -> Self {
        Self {
            remaining: maximum_facts,
        }
    }

    fn consume(&mut self, facts: usize) -> Result<(), ProjectionError> {
        self.remaining = self
            .remaining
            .checked_sub(facts)
            .ok_or(ProjectionError::TooManyStructuralFacts)?;
        Ok(())
    }
}

pub(super) fn materialize_structure(
    input: RawStructureInput<'_>,
    styles: Result<&StyleSheet, StructuralFactUnknownReason>,
    numbering: Result<&NumberingCatalog, StructuralFactUnknownReason>,
    maximum_facts: usize,
) -> Result<DocumentStructureFacts, ProjectionError> {
    let mut fact_budget = StructuralFactBudget::new(maximum_facts);
    let resolved = styles.and_then(|styles| {
        input
            .properties
            .iter()
            .map(|properties| styles.resolve(properties, numbering))
            .collect::<Result<Vec<_>, _>>()
    });

    let (indentation, numbering_hierarchy, outline_levels) = match resolved {
        Ok(properties) => {
            let indentation = properties
                .iter()
                .map(|properties| properties.indentation)
                .collect::<Result<Vec<_>, _>>()
                .map_or_else(StructuralFactSet::Unknown, |indentations| {
                    StructuralFactSet::Known(
                        indentations
                            .into_iter()
                            .enumerate()
                            .filter(|(_, indentation)| !indentation.is_empty())
                            .map(|(paragraph_ordinal, value)| ParagraphIndentationFact {
                                paragraph_ordinal,
                                value,
                            })
                            .collect(),
                    )
                });
            (
                indentation,
                materialize_numbering(&properties),
                StructuralFactSet::Known(
                    properties
                        .iter()
                        .enumerate()
                        .filter_map(|(paragraph_ordinal, properties)| {
                            properties.outline_level.map(|outline_level| {
                                ParagraphOutlineLevelFact {
                                    paragraph_ordinal,
                                    outline_level,
                                }
                            })
                        })
                        .collect(),
                ),
            )
        }
        Err(reason) => (
            StructuralFactSet::Unknown(reason),
            StructuralFactSet::Unknown(reason),
            StructuralFactSet::Unknown(reason),
        ),
    };
    if let StructuralFactSet::Known(facts) = &indentation {
        fact_budget.consume(facts.len())?;
    }
    if let StructuralFactSet::Known(facts) = &numbering_hierarchy {
        fact_budget.consume(facts.len())?;
    }
    if let StructuralFactSet::Known(facts) = &outline_levels {
        fact_budget.consume(facts.len())?;
    }

    let bookmarks = match input.bookmarks {
        Ok(ranges) => StructuralFactSet::Known(materialize_bookmarks(
            input.paragraph_texts,
            ranges,
            &mut fact_budget,
        )?),
        Err(reason) => StructuralFactSet::Unknown(reason),
    };
    let internal_references = match (input.references, &bookmarks) {
        (Ok([]), _) => StructuralFactSet::Known(Vec::new()),
        (Ok(references), StructuralFactSet::Known(bookmark_facts)) => StructuralFactSet::Known(
            materialize_references(references, bookmark_facts, &mut fact_budget)?,
        ),
        (Err(reason), _) => StructuralFactSet::Unknown(reason),
        (_, StructuralFactSet::Unknown(_)) => {
            StructuralFactSet::Unknown(StructuralFactUnknownReason::IncompleteBookmarkRanges)
        }
    };

    Ok(DocumentStructureFacts {
        indentation,
        numbering_hierarchy,
        bookmarks,
        internal_references,
        outline_levels,
    })
}

fn materialize_numbering(
    properties: &[ResolvedParagraphProperties],
) -> StructuralFactSet<NumberingHierarchyFact> {
    let mut stacks: HashMap<u32, [Option<usize>; 9]> = HashMap::new();
    let mut parents = vec![None; properties.len()];
    let mut children = vec![Vec::new(); properties.len()];

    for (ordinal, paragraph_properties) in properties.iter().enumerate() {
        if !paragraph_properties.numbering.present
            || paragraph_properties.numbering.num_id == Some(0)
        {
            continue;
        }
        let Some(num_id) = paragraph_properties.numbering.num_id else {
            return StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedNumbering);
        };
        let level = usize::from(paragraph_properties.numbering.level.unwrap_or(0));
        if level >= 9 {
            return StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedNumbering);
        }
        let stack = stacks.entry(num_id).or_insert([None; 9]);
        if level > 0
            && let Some(parent) = stack
                .get(..level)
                .and_then(|ancestors| ancestors.iter().rev().flatten().next().copied())
        {
            let Some(parent_slot) = parents.get_mut(ordinal) else {
                return StructuralFactSet::Unknown(
                    StructuralFactUnknownReason::UnsupportedNumbering,
                );
            };
            *parent_slot = Some(parent);
            let Some(parent_children) = children.get_mut(parent) else {
                return StructuralFactSet::Unknown(
                    StructuralFactUnknownReason::UnsupportedNumbering,
                );
            };
            parent_children.push(ordinal);
        }
        let Some(level_slot) = stack.get_mut(level) else {
            return StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedNumbering);
        };
        *level_slot = Some(ordinal);
        let descendant_start = level.saturating_add(1);
        let Some(descendant_slots) = stack.get_mut(descendant_start..) else {
            return StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedNumbering);
        };
        for slot in descendant_slots {
            *slot = None;
        }
    }

    StructuralFactSet::Known(
        properties
            .iter()
            .enumerate()
            .filter_map(|(ordinal, _)| {
                let parent = parents.get(ordinal).copied().flatten();
                let paragraph_children = children.get(ordinal)?;
                (parent.is_some() || !paragraph_children.is_empty()).then(|| {
                    NumberingHierarchyFact {
                        paragraph_ordinal: ordinal,
                        parent_paragraph_ordinal: parent,
                        child_paragraph_ordinals: paragraph_children.clone(),
                    }
                })
            })
            .collect(),
    )
}

fn materialize_bookmarks(
    texts: &[&str],
    ranges: &[RawBookmarkRange],
    fact_budget: &mut StructuralFactBudget,
) -> Result<Vec<BookmarkFact>, ProjectionError> {
    let mut facts = Vec::new();
    for range in ranges {
        segment_range(texts, &range.start, &range.end, |paragraph, span| {
            fact_budget.consume(1)?;
            facts.push(BookmarkFact {
                paragraph_ordinal: paragraph,
                bookmark_id: range.id,
                name: range.name.clone(),
                span,
            });
            Ok(())
        })?;
    }
    facts.sort_by(|left, right| {
        (
            left.paragraph_ordinal,
            left.span.start_utf8,
            left.span.end_utf8,
            left.bookmark_id,
            left.name.as_str(),
            left.span.coverage,
        )
            .cmp(&(
                right.paragraph_ordinal,
                right.span.start_utf8,
                right.span.end_utf8,
                right.bookmark_id,
                right.name.as_str(),
                right.span.coverage,
            ))
    });
    Ok(facts)
}

fn materialize_references(
    references: &[RawInternalReference],
    bookmarks: &[BookmarkFact],
    fact_budget: &mut StructuralFactBudget,
) -> Result<Vec<InternalReferenceFact>, ProjectionError> {
    let mut facts = Vec::new();
    let mut targets_by_name: HashMap<&str, Vec<&BookmarkFact>> = HashMap::new();
    for bookmark in bookmarks {
        targets_by_name
            .entry(&bookmark.name)
            .or_default()
            .push(bookmark);
    }
    let mut emitted_target_names = HashSet::new();
    for reference in references {
        fact_budget.consume(1)?;
        facts.push(InternalReferenceFact {
            paragraph_ordinal: reference.source.paragraph,
            reference_id: reference.reference_id.clone(),
            role: InternalReferenceRole::Source,
            span: StructuralSpan {
                start_utf8: reference.source.utf8,
                end_utf8: reference.source.utf8,
                start_utf16: reference.source.utf16,
                end_utf16: reference.source.utf16,
                coverage: SpanCoverage::Complete,
            },
        });
        if emitted_target_names.insert(reference.reference_id.as_str())
            && let Some(targets) = targets_by_name.get(reference.reference_id.as_str())
        {
            for bookmark in targets {
                fact_budget.consume(1)?;
                facts.push(InternalReferenceFact {
                    paragraph_ordinal: bookmark.paragraph_ordinal,
                    reference_id: reference.reference_id.clone(),
                    role: InternalReferenceRole::Target,
                    span: bookmark.span,
                });
            }
        }
    }
    facts.sort_by(|left, right| {
        (
            left.paragraph_ordinal,
            left.span.start_utf8,
            left.span.end_utf8,
            left.reference_id.as_str(),
            left.role,
            left.span.coverage,
        )
            .cmp(&(
                right.paragraph_ordinal,
                right.span.start_utf8,
                right.span.end_utf8,
                right.reference_id.as_str(),
                right.role,
                right.span.coverage,
            ))
    });
    Ok(facts)
}

fn segment_range(
    texts: &[&str],
    start: &RawBlockPoint,
    end: &RawBlockPoint,
    mut visit: impl FnMut(usize, StructuralSpan) -> Result<(), ProjectionError>,
) -> Result<(), ProjectionError> {
    if start.paragraph > end.paragraph || end.paragraph >= texts.len() {
        return Err(ProjectionError::InvalidDocumentXml);
    }
    if start.paragraph == end.paragraph && (start.utf8 > end.utf8 || start.utf16 > end.utf16) {
        return Err(ProjectionError::InvalidDocumentXml);
    }
    for (paragraph, text) in texts
        .iter()
        .enumerate()
        .take(end.paragraph.saturating_add(1))
        .skip(start.paragraph)
    {
        let text_utf8 =
            u32::try_from(text.len()).map_err(|_| ProjectionError::InvalidDocumentXml)?;
        let text_utf16 = u32::try_from(text.encode_utf16().count())
            .map_err(|_| ProjectionError::InvalidDocumentXml)?;
        let (start_utf8, start_utf16) = if paragraph == start.paragraph {
            (start.utf8, start.utf16)
        } else {
            (0, 0)
        };
        let (end_utf8, end_utf16) = if paragraph == end.paragraph {
            (end.utf8, end.utf16)
        } else {
            (text_utf8, text_utf16)
        };
        if start_utf8 > text_utf8
            || end_utf8 > text_utf8
            || start_utf16 > text_utf16
            || end_utf16 > text_utf16
        {
            return Err(ProjectionError::InvalidDocumentXml);
        }
        let continues_before = paragraph > start.paragraph;
        let continues_after = paragraph < end.paragraph;
        let coverage = match (continues_before, continues_after) {
            (false, false) => SpanCoverage::Complete,
            (true, false) => SpanCoverage::ContinuesBefore,
            (false, true) => SpanCoverage::ContinuesAfter,
            (true, true) => SpanCoverage::ContinuesBeforeAndAfter,
        };
        visit(
            paragraph,
            StructuralSpan {
                start_utf8,
                end_utf8,
                start_utf16,
                end_utf16,
                coverage,
            },
        )?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{
        ParagraphProperties, StyleDefinition, StyleKind, StyleSheet, TextProperties,
        inherit_style_toggle,
    };

    const WORD_STYLE_LIMIT: usize = 4_079;

    fn style(based_on: Option<String>, kind: StyleKind) -> StyleDefinition {
        StyleDefinition {
            based_on,
            properties: ParagraphProperties::default(),
            text_properties: TextProperties {
                bold: Some(true),
                ..TextProperties::default()
            },
            kind,
        }
    }

    #[test]
    fn style_toggle_truth_table_preserves_false_and_toggles_true() {
        let cases = [
            (&[][..], None),
            (&[false][..], None),
            (&[true][..], Some(true)),
            (&[true, false][..], Some(true)),
            (&[true, true][..], Some(false)),
            (&[true, false, true][..], Some(false)),
            (&[true, true, true][..], Some(true)),
        ];

        for (levels, expected) in cases {
            let actual = levels.iter().fold(None, |inherited, &level| {
                inherit_style_toggle(inherited, Some(level))
            });
            assert_eq!(actual, expected, "style levels {levels:?}");
        }
    }

    #[test]
    fn style_resolution_stays_within_its_linear_work_budget() -> Result<(), &'static str> {
        let ids = (0..WORD_STYLE_LIMIT)
            .map(|index| format!("style-{index}"))
            .collect::<Vec<_>>();

        let mut paragraph_chain = StyleSheet::default();
        let mut paragraph_parent = None;
        for id in &ids {
            paragraph_chain.styles.insert(
                id.clone(),
                style(paragraph_parent.clone(), StyleKind::Paragraph),
            );
            paragraph_parent = Some(id.clone());
        }
        let (resolved_paragraphs, paragraph_steps) = paragraph_chain.resolve_paragraph_styles();
        assert_eq!(resolved_paragraphs.len(), WORD_STYLE_LIMIT);
        assert!(paragraph_steps <= WORD_STYLE_LIMIT.saturating_mul(2).saturating_add(1));

        let mut character_chain = StyleSheet::default();
        let mut character_parent = None;
        for id in &ids {
            character_chain.styles.insert(
                id.clone(),
                style(character_parent.clone(), StyleKind::Character),
            );
            character_parent = Some(id.clone());
        }
        let (resolved_characters, character_steps) =
            character_chain.resolve_character_text_styles();
        assert_eq!(resolved_characters.len(), WORD_STYLE_LIMIT);
        assert!(character_steps <= WORD_STYLE_LIMIT.saturating_mul(2).saturating_add(1));

        let mut cycle = StyleSheet::default();
        let first = ids.first().ok_or("the fixture must be non-empty")?;
        for (index, id) in ids.iter().enumerate() {
            let successor = ids.get(index.saturating_add(1)).unwrap_or(first);
            cycle.styles.insert(
                id.clone(),
                style(Some(successor.clone()), StyleKind::Paragraph),
            );
        }
        let (cycle_resolved, cycle_steps) = cycle.resolve_paragraph_styles();
        assert!(cycle_resolved.is_empty());
        assert!(cycle_steps <= WORD_STYLE_LIMIT.saturating_add(1));
        Ok(())
    }
}
