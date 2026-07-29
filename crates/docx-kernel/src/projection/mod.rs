//! Bounded, host-independent projection of the main OOXML document part.
//!
//! This crate deliberately knows nothing about host-specific paragraph identifiers.
//! Callers allocate opaque application identities from package facts; `w14:paraId`
//! remains optional package metadata and is never treated as a host navigation ID.

mod archive;
mod compatibility;
mod namespaces;
mod numbering;
mod ooxml;
mod relationships;
mod review;
mod structure;
mod styles;

use std::collections::{HashMap, HashSet};
use std::fmt;

pub use archive::{DocumentParts, DocxLimits, extract_document_parts, extract_document_xml};
pub use ooxml::{
    PackageParagraphId, ParagraphStructure, RevisionProjectionStatus, RevisionUnsupportedReason,
    RevisionView, TextFormattingSpan, TextMaterialization, TextStyle,
};
pub use review::{
    AttributedComment, AttributedRevision, CommentContent, DocumentReviewFacts, ReviewDetail,
    ReviewFactLimits, ReviewFactSet, ReviewFactUnknownReason, ReviewPoint, ReviewSpan,
    RevisionContent, RevisionFactKind,
};
pub use structure::{
    BookmarkFact, DocumentStructureFacts, InternalReferenceFact, InternalReferenceRole,
    NumberingHierarchyFact, ParagraphIndentation, ParagraphIndentationFact,
    ParagraphOutlineLevelFact, SpanCoverage, StructuralFactSet, StructuralFactUnknownReason,
    StructuralSpan,
};
pub use styles::is_semantic_highlight_color;

#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub struct InternalParagraphId(String);

impl InternalParagraphId {
    /// Creates a bounded, non-empty application paragraph identity.
    ///
    /// # Errors
    ///
    /// Returns [`ProjectionError::InvalidInternalParagraphId`] when the value
    /// is empty or exceeds the identity length limit.
    pub fn new(value: impl Into<String>) -> Result<Self, ProjectionError> {
        let value = value.into();
        if value.is_empty() || value.len() > 128 {
            return Err(ProjectionError::InvalidInternalParagraphId);
        }
        Ok(Self(value))
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Clone, Copy, Debug)]
pub struct ParagraphIdentityFacts<'a> {
    pub ordinal: usize,
    pub package_paragraph_id: Option<PackageParagraphId>,
    pub text: &'a str,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ProjectedParagraph {
    pub id: InternalParagraphId,
    pub ordinal: usize,
    pub package_paragraph_id: Option<PackageParagraphId>,
    pub style_id: Option<String>,
    pub text: String,
    pub formatting: Vec<TextFormattingSpan>,
    pub structure: Option<ParagraphStructure>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum FormattingUnknownReason {
    DocumentPartOnly,
    StylesPartUnavailable,
    UnsupportedStyles,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum FormattingProjectionStatus {
    Complete,
    Incomplete(FormattingUnknownReason),
}

impl Default for FormattingProjectionStatus {
    fn default() -> Self {
        Self::Incomplete(FormattingUnknownReason::DocumentPartOnly)
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DocumentProjection {
    pub paragraphs: Vec<ProjectedParagraph>,
    pub formatting_status: FormattingProjectionStatus,
    pub revision_status: RevisionProjectionStatus,
    pub structural_facts: DocumentStructureFacts,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DocumentPackageProjection {
    pub document: DocumentProjection,
    pub review_facts: DocumentReviewFacts,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ProjectionOptions {
    pub revision_view: RevisionView,
    pub text_materialization: TextMaterialization,
}

impl Default for ProjectionOptions {
    fn default() -> Self {
        Self {
            revision_view: RevisionView::Current,
            text_materialization: TextMaterialization::WordHost,
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ProjectionError {
    ArchiveTooLarge,
    InvalidArchive,
    TooManyArchiveEntries,
    InvalidPackageRelationships,
    PackageRelationshipsTooLarge,
    MissingDocumentXml,
    DuplicateDocumentXml,
    DuplicateStylesXml,
    DuplicateNumberingXml,
    EncryptedDocumentXml,
    UnsupportedCompression(u16),
    DocumentXmlTooLarge,
    SuspiciousCompressionRatio,
    InvalidDocumentXmlEntry,
    DocumentXmlIntegrity,
    StylesXmlTooLarge,
    InvalidStylesXmlEntry,
    StylesXmlIntegrity,
    NumberingXmlTooLarge,
    InvalidNumberingXmlEntry,
    NumberingXmlIntegrity,
    InvalidDocumentXml,
    InvalidStylesXml,
    InvalidNumberingXml,
    MissingDocumentBody,
    TooManyParagraphs,
    TooManyStructuralFacts,
    TooManyNumberingItems,
    InvalidInternalParagraphId,
    DuplicateInternalParagraphId,
}

impl fmt::Display for ProjectionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let message = match self {
            Self::ArchiveTooLarge => "DOCX archive exceeds the configured size limit",
            Self::InvalidArchive => "DOCX archive is invalid",
            Self::TooManyArchiveEntries => "DOCX archive has too many entries",
            Self::InvalidPackageRelationships => "DOCX archive has invalid package relationships",
            Self::PackageRelationshipsTooLarge => {
                "DOCX package relationships exceed the configured size limit"
            }
            Self::MissingDocumentXml => "DOCX archive has no main document part",
            Self::DuplicateDocumentXml => "DOCX archive has duplicate main document parts",
            Self::DuplicateStylesXml => "DOCX archive has duplicate entries for its styles part",
            Self::DuplicateNumberingXml => {
                "DOCX archive has duplicate entries for its numbering part"
            }
            Self::EncryptedDocumentXml => "DOCX main document part is encrypted",
            Self::UnsupportedCompression(_) => {
                "selected DOCX package part uses unsupported compression"
            }
            Self::DocumentXmlTooLarge => {
                "DOCX main document part exceeds the configured size limit"
            }
            Self::SuspiciousCompressionRatio => {
                "selected DOCX package part exceeds the compression-ratio limit"
            }
            Self::InvalidDocumentXmlEntry => "DOCX main document part has an invalid ZIP entry",
            Self::DocumentXmlIntegrity => "DOCX main document part failed size or CRC validation",
            Self::StylesXmlTooLarge => "DOCX styles part exceeds the configured size limit",
            Self::InvalidStylesXmlEntry => "DOCX styles part has an invalid ZIP entry",
            Self::StylesXmlIntegrity => "DOCX styles part failed size or CRC validation",
            Self::NumberingXmlTooLarge => "DOCX numbering part exceeds the configured size limit",
            Self::InvalidNumberingXmlEntry => "DOCX numbering part has an invalid ZIP entry",
            Self::NumberingXmlIntegrity => "DOCX numbering part failed size or CRC validation",
            Self::InvalidDocumentXml => "DOCX main document part is invalid XML",
            Self::InvalidStylesXml => "DOCX styles part is invalid XML",
            Self::InvalidNumberingXml => "DOCX numbering part is invalid XML",
            Self::MissingDocumentBody => "DOCX main document part has no document body",
            Self::TooManyParagraphs => "DOCX main document part has too many paragraphs",
            Self::TooManyStructuralFacts => {
                "DOCX main document part produces too many structural facts"
            }
            Self::TooManyNumberingItems => "DOCX numbering part exceeds the configured item limit",
            Self::InvalidInternalParagraphId => "application paragraph ID is invalid",
            Self::DuplicateInternalParagraphId => "application paragraph IDs are not unique",
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for ProjectionError {}

/// Projects a bounded DOCX package using default projection options.
///
/// # Errors
///
/// Returns [`ProjectionError`] for invalid or unsupported package input, a
/// resource-limit violation, or an invalid identity allocated by the caller.
pub fn project_docx<F>(
    bytes: &[u8],
    limits: DocxLimits,
    allocate_id: F,
) -> Result<DocumentProjection, ProjectionError>
where
    F: FnMut(ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError>,
{
    project_docx_with_options(bytes, limits, ProjectionOptions::default(), allocate_id)
}

/// Projects a bounded DOCX package using explicit projection options.
///
/// # Errors
///
/// Returns [`ProjectionError`] for invalid or unsupported package input, a
/// resource-limit violation, or an invalid identity allocated by the caller.
pub fn project_docx_with_options<F>(
    bytes: &[u8],
    limits: DocxLimits,
    options: ProjectionOptions,
    allocate_id: F,
) -> Result<DocumentProjection, ProjectionError>
where
    F: FnMut(ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError>,
{
    let parts = extract_document_parts(bytes, limits)?;
    let styles = parts.styles_xml.as_deref().map_or(
        Err(StructuralFactUnknownReason::StylesPartUnavailable),
        |styles| {
            styles::parse_styles(styles, limits.maximum_styles)
                .map_err(|_| StructuralFactUnknownReason::UnsupportedStyles)
        },
    );
    let numbering = parse_optional_numbering(
        parts.numbering_xml.as_deref(),
        limits.maximum_numbering_items,
    )?;
    project_document_xml_with_limit(
        &parts.document_xml,
        limits.maximum_paragraphs,
        limits.maximum_structural_facts,
        options,
        ProjectionDependencies {
            styles: styles.as_ref().map_err(|reason| *reason),
            numbering: numbering
                .as_ref()
                .ok_or(StructuralFactUnknownReason::UnsupportedNumbering),
        },
        allocate_id,
    )
}

/// Projects a bounded DOCX package and its attributed review facts in one
/// package-directory scan.
///
/// Invalid optional review parts produce an explicit unknown fact family; they
/// do not discard an otherwise valid document projection.
///
/// # Errors
///
/// Returns [`ProjectionError`] for an invalid document projection or package
/// boundary. Optional comments-part failures remain represented in
/// [`DocumentReviewFacts`].
pub fn project_docx_with_review_facts<F>(
    bytes: &[u8],
    limits: DocxLimits,
    review_limits: ReviewFactLimits,
    options: ProjectionOptions,
    allocate_id: F,
) -> Result<DocumentPackageProjection, ProjectionError>
where
    F: FnMut(ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError>,
{
    let parts = archive::extract_projection_parts(bytes, limits, review_limits)?;
    let styles = parts.styles.as_deref().map_or(
        Err(StructuralFactUnknownReason::StylesPartUnavailable),
        |styles| {
            styles::parse_styles(styles, limits.maximum_styles)
                .map_err(|_| StructuralFactUnknownReason::UnsupportedStyles)
        },
    );
    let numbering =
        parse_optional_numbering(parts.numbering.as_deref(), limits.maximum_numbering_items)?;
    let ProjectedDocumentWithReview {
        document,
        revisions,
        comment_anchors,
    } = project_document_xml_with_limit_and_review(
        &parts.document,
        limits.maximum_paragraphs,
        limits.maximum_structural_facts,
        options,
        ProjectionDependencies {
            styles: styles.as_ref().map_err(|reason| *reason),
            numbering: numbering
                .as_ref()
                .ok_or(StructuralFactUnknownReason::UnsupportedNumbering),
        },
        Some(ooxml::ReviewProjectionLimits {
            maximum_facts: review_limits.maximum_facts_per_family,
            maximum_detail_bytes: review_limits.maximum_review_detail_bytes,
        }),
        allocate_id,
    )?;
    let review_facts = review::project_review_facts(
        revisions.unwrap_or(ReviewFactSet::Unknown(
            ReviewFactUnknownReason::InvalidDocument,
        )),
        comment_anchors.as_ref(),
        &document,
        parts.comments,
        parts.comments_extended,
        review_limits,
        options.text_materialization,
    );
    Ok(DocumentPackageProjection {
        document,
        review_facts,
    })
}

fn parse_optional_numbering(
    xml: Option<&[u8]>,
    maximum_items: usize,
) -> Result<Option<numbering::NumberingCatalog>, ProjectionError> {
    let Some(xml) = xml else {
        return Ok(None);
    };
    match numbering::parse_numbering(xml, maximum_items) {
        Ok(catalog) => Ok(Some(catalog)),
        Err(ProjectionError::TooManyNumberingItems) => Err(ProjectionError::TooManyNumberingItems),
        Err(_) => Ok(None),
    }
}

/// Projects an uncompressed main OOXML document part with default options.
///
/// # Errors
///
/// Returns [`ProjectionError`] for invalid or unsupported XML, a resource-limit
/// violation, or an invalid identity allocated by the caller.
pub fn project_document_xml<F>(
    xml: &[u8],
    allocate_id: F,
) -> Result<DocumentProjection, ProjectionError>
where
    F: FnMut(ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError>,
{
    project_document_xml_with_options(xml, ProjectionOptions::default(), allocate_id)
}

/// Projects an uncompressed main OOXML document part with explicit options.
///
/// # Errors
///
/// Returns [`ProjectionError`] for invalid or unsupported XML, a resource-limit
/// violation, or an invalid identity allocated by the caller.
pub fn project_document_xml_with_options<F>(
    xml: &[u8],
    options: ProjectionOptions,
    allocate_id: F,
) -> Result<DocumentProjection, ProjectionError>
where
    F: FnMut(ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError>,
{
    project_document_xml_with_limit(
        xml,
        DocxLimits::default().maximum_paragraphs,
        DocxLimits::default().maximum_structural_facts,
        options,
        ProjectionDependencies {
            styles: Err(StructuralFactUnknownReason::DocumentPartOnly),
            numbering: Err(StructuralFactUnknownReason::DocumentPartOnly),
        },
        allocate_id,
    )
}

fn project_document_xml_with_limit<F>(
    xml: &[u8],
    maximum_paragraphs: usize,
    maximum_structural_facts: usize,
    options: ProjectionOptions,
    dependencies: ProjectionDependencies<'_>,
    allocate_id: F,
) -> Result<DocumentProjection, ProjectionError>
where
    F: FnMut(ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError>,
{
    project_document_xml_with_limit_and_review(
        xml,
        maximum_paragraphs,
        maximum_structural_facts,
        options,
        dependencies,
        None,
        allocate_id,
    )
    .map(|projection| projection.document)
}

struct ProjectedDocumentWithReview {
    document: DocumentProjection,
    revisions: Option<ReviewFactSet<AttributedRevision>>,
    comment_anchors: Option<HashMap<String, ReviewSpan>>,
}

#[derive(Clone, Copy)]
struct ProjectionDependencies<'a> {
    styles: Result<&'a structure::StyleSheet, StructuralFactUnknownReason>,
    numbering: Result<&'a numbering::NumberingCatalog, StructuralFactUnknownReason>,
}

fn project_document_xml_with_limit_and_review<F>(
    xml: &[u8],
    maximum_paragraphs: usize,
    maximum_structural_facts: usize,
    options: ProjectionOptions,
    dependencies: ProjectionDependencies<'_>,
    review_limits: Option<ooxml::ReviewProjectionLimits>,
    mut allocate_id: F,
) -> Result<ProjectedDocumentWithReview, ProjectionError>
where
    F: FnMut(ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError>,
{
    let projected = ooxml::project_document_xml(
        xml,
        maximum_paragraphs,
        options.revision_view,
        options.text_materialization,
        dependencies.styles.map_err(formatting_unknown_reason),
        review_limits,
    )?;
    let review_revisions = projected.review_revisions;
    let review_comment_anchors = review_limits
        .is_some()
        .then_some(projected.review_comment_anchors);
    let mut seen_ids = HashSet::with_capacity(projected.paragraphs.len());
    let mut ids = Vec::with_capacity(projected.paragraphs.len());
    for paragraph in &projected.paragraphs {
        let facts = ParagraphIdentityFacts {
            ordinal: paragraph.ordinal,
            package_paragraph_id: paragraph.package_paragraph_id,
            text: &paragraph.text,
        };
        let id = allocate_id(facts)?;
        if !seen_ids.insert(id.clone()) {
            return Err(ProjectionError::DuplicateInternalParagraphId);
        }
        ids.push(id);
    }
    let texts = projected
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.text.as_str())
        .collect::<Vec<_>>();
    let properties = projected
        .paragraphs
        .iter()
        .map(|paragraph| &paragraph.properties)
        .collect::<Vec<_>>();
    let structural_facts = structure::materialize_structure(
        structure::RawStructureInput {
            paragraph_texts: &texts,
            properties: &properties,
            bookmarks: projected.bookmarks.as_deref().map_err(|reason| *reason),
            references: projected.references.as_deref().map_err(|reason| *reason),
        },
        dependencies.styles,
        dependencies.numbering,
        maximum_structural_facts,
    )?;
    let mut paragraphs = Vec::with_capacity(projected.paragraphs.len());
    for (id, paragraph) in ids.into_iter().zip(projected.paragraphs) {
        paragraphs.push(ProjectedParagraph {
            id,
            ordinal: paragraph.ordinal,
            package_paragraph_id: paragraph.package_paragraph_id,
            style_id: paragraph.properties.style_id,
            text: paragraph.text,
            formatting: paragraph.formatting,
            structure: paragraph.structure,
        });
    }
    Ok(ProjectedDocumentWithReview {
        document: DocumentProjection {
            paragraphs,
            formatting_status: projected.formatting_status,
            revision_status: projected.revision_status,
            structural_facts,
        },
        revisions: review_revisions,
        comment_anchors: review_comment_anchors,
    })
}

const fn formatting_unknown_reason(reason: StructuralFactUnknownReason) -> FormattingUnknownReason {
    match reason {
        StructuralFactUnknownReason::DocumentPartOnly => FormattingUnknownReason::DocumentPartOnly,
        StructuralFactUnknownReason::StylesPartUnavailable => {
            FormattingUnknownReason::StylesPartUnavailable
        }
        StructuralFactUnknownReason::UnsupportedStyles
        | StructuralFactUnknownReason::UnsupportedNumbering
        | StructuralFactUnknownReason::IncompleteBookmarkRanges
        | StructuralFactUnknownReason::UnsupportedInternalReferences => {
            FormattingUnknownReason::UnsupportedStyles
        }
    }
}
