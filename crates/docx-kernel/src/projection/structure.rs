use std::collections::{HashMap, HashSet};

use crate::ProjectionError;

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
    pub(super) fn is_empty(self) -> bool {
        self == Self::default()
    }

    pub(super) fn inherit(self, child: Self) -> Self {
        Self {
            first_line_twips: if child.hanging_twips.is_some() {
                None
            } else {
                child.first_line_twips.or(self.first_line_twips)
            },
            hanging_twips: if child.first_line_twips.is_some() {
                None
            } else {
                child.hanging_twips.or(self.hanging_twips)
            },
            left_twips: child.left_twips.or(self.left_twips),
            right_twips: child.right_twips.or(self.right_twips),
            start_twips: child.start_twips.or(self.start_twips),
            end_twips: child.end_twips.or(self.end_twips),
            first_line_chars_hundredths: if child.hanging_chars_hundredths.is_some() {
                None
            } else {
                child
                    .first_line_chars_hundredths
                    .or(self.first_line_chars_hundredths)
            },
            hanging_chars_hundredths: if child.first_line_chars_hundredths.is_some() {
                None
            } else {
                child
                    .hanging_chars_hundredths
                    .or(self.hanging_chars_hundredths)
            },
            left_chars_hundredths: child.left_chars_hundredths.or(self.left_chars_hundredths),
            right_chars_hundredths: child.right_chars_hundredths.or(self.right_chars_hundredths),
            start_chars_hundredths: child.start_chars_hundredths.or(self.start_chars_hundredths),
            end_chars_hundredths: child.end_chars_hundredths.or(self.end_chars_hundredths),
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ParagraphIndentationFact {
    pub paragraph_ordinal: usize,
    pub value: ParagraphIndentation,
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
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub(super) struct NumberingProperties {
    pub num_id: Option<u32>,
    pub level: Option<u8>,
    pub present: bool,
}

impl NumberingProperties {
    pub(super) fn inherit(self, child: Self) -> Self {
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
    pub indentation: ParagraphIndentation,
    pub numbering: NumberingProperties,
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
}

#[derive(Clone, Debug, Default)]
pub(super) struct StyleSheet {
    pub default_style_id: Option<String>,
    pub document_defaults: ParagraphProperties,
    pub styles: HashMap<String, StyleDefinition>,
}

impl StyleSheet {
    pub(super) fn resolve(
        &self,
        direct: &ParagraphProperties,
    ) -> Result<ParagraphProperties, StructuralFactUnknownReason> {
        let initial_style_id = direct.style_id.as_ref().or(self.default_style_id.as_ref());
        let mut chain = Vec::new();
        let mut current = initial_style_id;
        while let Some(current_style_id) = current {
            if chain.iter().any(|seen| seen == current_style_id) || chain.len() >= 128 {
                return Err(StructuralFactUnknownReason::UnsupportedStyles);
            }
            let Some(style) = self.styles.get(current_style_id) else {
                return Err(StructuralFactUnknownReason::UnsupportedStyles);
            };
            chain.push(current_style_id.clone());
            current = style.based_on.as_ref();
        }

        let mut resolved = self.document_defaults.clone();
        for current_style_id in chain.iter().rev() {
            let style = self
                .styles
                .get(current_style_id)
                .ok_or(StructuralFactUnknownReason::UnsupportedStyles)?;
            resolved.indentation = resolved.indentation.inherit(style.properties.indentation);
            resolved.numbering = resolved.numbering.inherit(style.properties.numbering);
        }
        resolved.indentation = resolved.indentation.inherit(direct.indentation);
        resolved.numbering = resolved.numbering.inherit(direct.numbering);
        Ok(resolved)
    }
}

#[derive(Clone, Copy)]
pub(super) struct RawStructureInput<'a> {
    pub paragraph_texts: &'a [&'a str],
    pub properties: &'a [ParagraphProperties],
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
    maximum_facts: usize,
) -> Result<DocumentStructureFacts, ProjectionError> {
    let mut fact_budget = StructuralFactBudget::new(maximum_facts);
    let resolved = styles.and_then(|styles| {
        input
            .properties
            .iter()
            .map(|properties| styles.resolve(properties))
            .collect::<Result<Vec<_>, _>>()
    });

    let (indentation, numbering_hierarchy) = match resolved {
        Ok(properties) => (
            StructuralFactSet::Known(
                properties
                    .iter()
                    .enumerate()
                    .filter(|(_, properties)| !properties.indentation.is_empty())
                    .map(|(ordinal, properties)| ParagraphIndentationFact {
                        paragraph_ordinal: ordinal,
                        value: properties.indentation,
                    })
                    .collect(),
            ),
            materialize_numbering(&properties),
        ),
        Err(reason) => (
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
    })
}

fn materialize_numbering(
    properties: &[ParagraphProperties],
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
