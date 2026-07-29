use miniz_oxide::inflate::{TINFLStatus, decompress_to_vec_with_limit};
use rawzip::{CompressionMethod, ZipArchive, ZipSliceArchive, crc32};

use crate::ProjectionError;
use crate::projection::relationships::{
    InvalidReviewRelationships, ReviewPartPaths, document_relationship_paths,
    document_relationships_path, main_document_path,
};
use crate::projection::review::{ReviewFactLimits, ReviewFactUnknownReason};

const DOCUMENT_XML_PATH: &[u8] = b"word/document.xml";
const ROOT_RELATIONSHIPS_PATH: &[u8] = b"_rels/.rels";
const MAXIMUM_ROOT_RELATIONSHIPS_BYTES: usize = 1024 * 1024;
const MAXIMUM_DOCUMENT_RELATIONSHIPS_BYTES: usize = 1024 * 1024;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DocxLimits {
    pub maximum_archive_bytes: usize,
    pub maximum_document_xml_bytes: usize,
    pub maximum_styles_xml_bytes: usize,
    pub maximum_styles: usize,
    pub maximum_numbering_xml_bytes: usize,
    pub maximum_numbering_items: usize,
    pub maximum_entries: u64,
    pub maximum_compression_ratio: u64,
    pub maximum_paragraphs: usize,
    pub maximum_structural_facts: usize,
}

impl Default for DocxLimits {
    fn default() -> Self {
        Self {
            maximum_archive_bytes: 64 * 1024 * 1024,
            maximum_document_xml_bytes: 32 * 1024 * 1024,
            maximum_styles_xml_bytes: 8 * 1024 * 1024,
            maximum_styles: 4_079,
            maximum_numbering_xml_bytes: 8 * 1024 * 1024,
            maximum_numbering_items: 100_000,
            maximum_entries: 4096,
            maximum_compression_ratio: 200,
            maximum_paragraphs: 250_000,
            maximum_structural_facts: 1_000_000,
        }
    }
}

#[derive(Clone, Copy)]
struct XmlEntry {
    wayfinder: rawzip::ZipArchiveEntryWayfinder,
    method: CompressionMethod,
    compressed_size: u64,
    uncompressed_size: u64,
    crc32: u32,
    encrypted: bool,
}

struct PackageEntry {
    path: Vec<u8>,
    xml: XmlEntry,
}

struct EntryExtractionErrors {
    invalid_entry: ProjectionError,
    too_large: ProjectionError,
    integrity: ProjectionError,
}

struct RelatedPartExtraction<'a> {
    path: Option<&'a [u8]>,
    maximum_bytes: usize,
    duplicate: ProjectionError,
    invalid_entry: ProjectionError,
    too_large: ProjectionError,
    errors: EntryExtractionErrors,
}

impl EntryExtractionErrors {
    const DOCUMENT_XML: Self = Self {
        invalid_entry: ProjectionError::InvalidDocumentXmlEntry,
        too_large: ProjectionError::DocumentXmlTooLarge,
        integrity: ProjectionError::DocumentXmlIntegrity,
    };
    const PACKAGE_RELATIONSHIPS: Self = Self {
        invalid_entry: ProjectionError::InvalidPackageRelationships,
        too_large: ProjectionError::PackageRelationshipsTooLarge,
        integrity: ProjectionError::InvalidPackageRelationships,
    };
    const STYLES_XML: Self = Self {
        invalid_entry: ProjectionError::InvalidStylesXmlEntry,
        too_large: ProjectionError::StylesXmlTooLarge,
        integrity: ProjectionError::StylesXmlIntegrity,
    };
    const NUMBERING_XML: Self = Self {
        invalid_entry: ProjectionError::InvalidNumberingXmlEntry,
        too_large: ProjectionError::NumberingXmlTooLarge,
        integrity: ProjectionError::NumberingXmlIntegrity,
    };
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DocumentParts {
    pub document_xml: Vec<u8>,
    pub styles_xml: Option<Vec<u8>>,
    pub numbering_xml: Option<Vec<u8>>,
}

pub(super) struct ProjectionPackageParts {
    pub document: Vec<u8>,
    pub styles: Option<Vec<u8>>,
    pub numbering: Option<Vec<u8>>,
    pub comments: Result<Option<Vec<u8>>, ReviewFactUnknownReason>,
    pub comments_extended: Result<Option<Vec<u8>>, ReviewFactUnknownReason>,
}

struct RequiredPackageParts {
    document: Vec<u8>,
    review_paths: Result<ReviewPartPaths, InvalidReviewRelationships>,
    styles: Option<Vec<u8>>,
    numbering: Option<Vec<u8>>,
}

struct ReviewPackageParts {
    comments: Result<Option<Vec<u8>>, ReviewFactUnknownReason>,
    comments_extended: Result<Option<Vec<u8>>, ReviewFactUnknownReason>,
}

/// Extracts the relationship-resolved main document part from a bounded DOCX package.
///
/// # Errors
///
/// Returns [`ProjectionError`] when the package is invalid, unsupported, or
/// exceeds the supplied resource limits.
pub fn extract_document_xml(bytes: &[u8], limits: DocxLimits) -> Result<Vec<u8>, ProjectionError> {
    extract_document_parts(bytes, limits).map(|parts| parts.document_xml)
}

/// Extracts the main document and its relationship-selected paragraph-style and
/// numbering parts from one bounded scan of the ZIP directory.
///
/// # Errors
///
/// Returns [`ProjectionError`] when the package is invalid, unsupported, or
/// exceeds the supplied resource limits.
pub fn extract_document_parts(
    bytes: &[u8],
    limits: DocxLimits,
) -> Result<DocumentParts, ProjectionError> {
    extract_required_parts(bytes, limits)
}

pub(super) fn extract_projection_parts(
    bytes: &[u8],
    limits: DocxLimits,
    review_limits: ReviewFactLimits,
) -> Result<ProjectionPackageParts, ProjectionError> {
    if bytes.len() > limits.maximum_archive_bytes {
        return Err(ProjectionError::ArchiveTooLarge);
    }
    let archive = ZipArchive::from_slice(bytes).map_err(|_| ProjectionError::InvalidArchive)?;
    let package_entries = scan_package_entries(&archive, limits)?;
    let required = extract_required_parts_from_archive(&archive, &package_entries, limits)?;
    let review = extract_review_parts(
        &archive,
        &package_entries,
        required.review_paths,
        review_limits,
        limits,
    );
    Ok(ProjectionPackageParts {
        document: required.document,
        styles: required.styles,
        numbering: required.numbering,
        comments: review.comments,
        comments_extended: review.comments_extended,
    })
}

fn extract_required_parts(
    bytes: &[u8],
    limits: DocxLimits,
) -> Result<DocumentParts, ProjectionError> {
    if bytes.len() > limits.maximum_archive_bytes {
        return Err(ProjectionError::ArchiveTooLarge);
    }
    let archive = ZipArchive::from_slice(bytes).map_err(|_| ProjectionError::InvalidArchive)?;
    let package_entries = scan_package_entries(&archive, limits)?;

    let required = extract_required_parts_from_archive(&archive, &package_entries, limits)?;
    Ok(DocumentParts {
        document_xml: required.document,
        styles_xml: required.styles,
        numbering_xml: required.numbering,
    })
}

fn extract_required_parts_from_archive(
    archive: &ZipSliceArchive<&[u8]>,
    package_entries: &[PackageEntry],
    limits: DocxLimits,
) -> Result<RequiredPackageParts, ProjectionError> {
    let document_path = resolve_document_path(archive, package_entries, limits)?;
    let document = unique_entry(
        package_entries,
        &document_path,
        ProjectionError::DuplicateDocumentXml,
    )?
    .ok_or(ProjectionError::MissingDocumentXml)?;
    validate_selected_entry(
        document,
        limits.maximum_document_xml_bytes,
        ProjectionError::DocumentXmlTooLarge,
        ProjectionError::EncryptedDocumentXml,
        limits,
    )?;
    let document_xml = extract_entry(
        archive,
        document,
        &document_path,
        limits.maximum_document_xml_bytes,
        EntryExtractionErrors::DOCUMENT_XML,
    )?;
    let relationships_path = document_relationships_path(&document_path)?;
    let document_relationships = unique_entry(
        package_entries,
        &relationships_path,
        ProjectionError::InvalidPackageRelationships,
    )?
    .map(|relationships| {
        validate_selected_entry(
            relationships,
            MAXIMUM_DOCUMENT_RELATIONSHIPS_BYTES,
            ProjectionError::PackageRelationshipsTooLarge,
            ProjectionError::InvalidPackageRelationships,
            limits,
        )?;
        extract_entry(
            archive,
            relationships,
            &relationships_path,
            MAXIMUM_DOCUMENT_RELATIONSHIPS_BYTES,
            EntryExtractionErrors::PACKAGE_RELATIONSHIPS,
        )
    })
    .transpose()?;
    let related_paths = document_relationships
        .as_deref()
        .map(|relationships| document_relationship_paths(relationships, &document_path))
        .transpose()?
        .unwrap_or_default();
    let styles_xml = extract_related_part(
        archive,
        package_entries,
        limits,
        RelatedPartExtraction {
            path: related_paths.styles.as_deref(),
            maximum_bytes: limits.maximum_styles_xml_bytes,
            duplicate: ProjectionError::DuplicateStylesXml,
            invalid_entry: ProjectionError::InvalidStylesXmlEntry,
            too_large: ProjectionError::StylesXmlTooLarge,
            errors: EntryExtractionErrors::STYLES_XML,
        },
    )?;
    let numbering_xml = extract_related_part(
        archive,
        package_entries,
        limits,
        RelatedPartExtraction {
            path: related_paths.numbering.as_deref(),
            maximum_bytes: limits.maximum_numbering_xml_bytes,
            duplicate: ProjectionError::DuplicateNumberingXml,
            invalid_entry: ProjectionError::InvalidNumberingXmlEntry,
            too_large: ProjectionError::NumberingXmlTooLarge,
            errors: EntryExtractionErrors::NUMBERING_XML,
        },
    )?;
    Ok(RequiredPackageParts {
        document: document_xml,
        review_paths: related_paths.review,
        styles: styles_xml,
        numbering: numbering_xml,
    })
}

fn extract_related_part(
    archive: &ZipSliceArchive<&[u8]>,
    entries: &[PackageEntry],
    limits: DocxLimits,
    extraction: RelatedPartExtraction<'_>,
) -> Result<Option<Vec<u8>>, ProjectionError> {
    let Some(path) = extraction.path else {
        return Ok(None);
    };
    let selected = unique_entry(entries, path, extraction.duplicate)?
        .ok_or_else(|| extraction.invalid_entry.clone())?;
    validate_selected_entry(
        selected,
        extraction.maximum_bytes,
        extraction.too_large,
        extraction.invalid_entry,
        limits,
    )?;
    extract_entry(
        archive,
        selected,
        path,
        extraction.maximum_bytes,
        extraction.errors,
    )
    .map(Some)
}

fn extract_review_parts(
    archive: &ZipSliceArchive<&[u8]>,
    entries: &[PackageEntry],
    paths: Result<ReviewPartPaths, InvalidReviewRelationships>,
    review_limits: ReviewFactLimits,
    limits: DocxLimits,
) -> ReviewPackageParts {
    let Ok(paths) = paths else {
        return ReviewPackageParts {
            comments: Err(ReviewFactUnknownReason::InvalidComments),
            comments_extended: Err(ReviewFactUnknownReason::InvalidCommentsExtended),
        };
    };
    let comments = paths.comments.map_or(Ok(None), |path| {
        extract_optional_review_entry(
            archive,
            entries,
            &path,
            review_limits.maximum_comments_xml_bytes,
            ReviewFactUnknownReason::InvalidComments,
            limits,
        )
        .and_then(|value| {
            value
                .ok_or(ReviewFactUnknownReason::InvalidComments)
                .map(Some)
        })
    });
    let comments_extended = paths.comments_extended.map_or(Ok(None), |path| {
        extract_optional_review_entry(
            archive,
            entries,
            &path,
            review_limits.maximum_comments_extended_xml_bytes,
            ReviewFactUnknownReason::InvalidCommentsExtended,
            limits,
        )
        .and_then(|value| {
            value
                .ok_or(ReviewFactUnknownReason::InvalidCommentsExtended)
                .map(Some)
        })
    });
    ReviewPackageParts {
        comments,
        comments_extended,
    }
}

fn extract_optional_review_entry(
    archive: &ZipSliceArchive<&[u8]>,
    entries: &[PackageEntry],
    path: &[u8],
    maximum_bytes: usize,
    invalid: ReviewFactUnknownReason,
    limits: DocxLimits,
) -> Result<Option<Vec<u8>>, ReviewFactUnknownReason> {
    let selected =
        unique_entry(entries, path, ProjectionError::InvalidArchive).map_err(|_| invalid)?;
    let Some(selected) = selected else {
        return Ok(None);
    };
    let validation = validate_selected_entry(
        selected,
        maximum_bytes,
        ProjectionError::DocumentXmlTooLarge,
        ProjectionError::InvalidDocumentXmlEntry,
        limits,
    );
    validation.map_err(|error| match error {
        ProjectionError::DocumentXmlTooLarge | ProjectionError::SuspiciousCompressionRatio => {
            ReviewFactUnknownReason::ResourceLimit
        }
        _ => invalid,
    })?;
    let extracted = extract_entry(
        archive,
        selected,
        path,
        maximum_bytes,
        EntryExtractionErrors::DOCUMENT_XML,
    );
    extracted.map(Some).map_err(|error| match error {
        ProjectionError::DocumentXmlTooLarge | ProjectionError::SuspiciousCompressionRatio => {
            ReviewFactUnknownReason::ResourceLimit
        }
        _ => invalid,
    })
}

fn scan_package_entries(
    archive: &ZipSliceArchive<&[u8]>,
    limits: DocxLimits,
) -> Result<Vec<PackageEntry>, ProjectionError> {
    if archive.entries_hint() > limits.maximum_entries {
        return Err(ProjectionError::TooManyArchiveEntries);
    }

    let mut package_entries = Vec::new();
    let mut entries = archive.entries();
    let mut entry_count = 0_u64;
    while let Some(entry) = entries
        .next_entry()
        .map_err(|_| ProjectionError::InvalidArchive)?
    {
        entry_count = entry_count
            .checked_add(1)
            .ok_or(ProjectionError::TooManyArchiveEntries)?;
        if entry_count > limits.maximum_entries {
            return Err(ProjectionError::TooManyArchiveEntries);
        }
        package_entries.push(PackageEntry {
            path: entry.file_path().as_ref().to_vec(),
            xml: XmlEntry {
                wayfinder: entry.wayfinder(),
                method: entry.compression_method(),
                compressed_size: entry.compressed_size_hint(),
                uncompressed_size: entry.uncompressed_size_hint(),
                crc32: entry.crc32(),
                encrypted: entry.flags().is_encrypted(),
            },
        });
    }
    Ok(package_entries)
}

fn resolve_document_path(
    archive: &ZipSliceArchive<&[u8]>,
    package_entries: &[PackageEntry],
    limits: DocxLimits,
) -> Result<Vec<u8>, ProjectionError> {
    if let Some(relationships) = unique_entry(
        package_entries,
        ROOT_RELATIONSHIPS_PATH,
        ProjectionError::InvalidPackageRelationships,
    )? {
        validate_selected_entry(
            relationships,
            MAXIMUM_ROOT_RELATIONSHIPS_BYTES,
            ProjectionError::PackageRelationshipsTooLarge,
            ProjectionError::InvalidPackageRelationships,
            limits,
        )?;
        let relationships_xml = extract_entry(
            archive,
            relationships,
            ROOT_RELATIONSHIPS_PATH,
            MAXIMUM_ROOT_RELATIONSHIPS_BYTES,
            EntryExtractionErrors::PACKAGE_RELATIONSHIPS,
        )?;
        main_document_path(&relationships_xml)
    } else {
        Ok(DOCUMENT_XML_PATH.to_vec())
    }
}

fn unique_entry(
    entries: &[PackageEntry],
    path: &[u8],
    duplicate: ProjectionError,
) -> Result<Option<XmlEntry>, ProjectionError> {
    let mut matches = entries
        .iter()
        .filter(|entry| entry.path == path)
        .map(|entry| entry.xml);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(duplicate);
    }
    Ok(first)
}

fn validate_selected_entry(
    entry: XmlEntry,
    maximum_bytes: usize,
    too_large: ProjectionError,
    encrypted: ProjectionError,
    limits: DocxLimits,
) -> Result<(), ProjectionError> {
    if entry.encrypted {
        return Err(encrypted);
    }
    if entry.method != CompressionMethod::STORE && entry.method != CompressionMethod::DEFLATE {
        return Err(ProjectionError::UnsupportedCompression(
            entry.method.as_u16(),
        ));
    }
    validate_declared_sizes(
        entry.compressed_size,
        entry.uncompressed_size,
        maximum_bytes,
        too_large,
        limits,
    )
}

fn extract_entry(
    archive: &ZipSliceArchive<&[u8]>,
    entry: XmlEntry,
    expected_path: &[u8],
    maximum_bytes: usize,
    errors: EntryExtractionErrors,
) -> Result<Vec<u8>, ProjectionError> {
    let local = archive
        .get_entry(entry.wayfinder)
        .map_err(|_| errors.invalid_entry.clone())?;
    let local_header = local.local_header();
    if local_header.compression_method() != entry.method
        || local_header.file_path().as_ref() != expected_path
        || local_header.flags().is_encrypted()
    {
        return Err(errors.invalid_entry);
    }
    let output = match entry.method {
        CompressionMethod::STORE => local.data().to_vec(),
        CompressionMethod::DEFLATE => decompress_to_vec_with_limit(local.data(), maximum_bytes)
            .map_err(|error| {
                if error.status == TINFLStatus::HasMoreOutput {
                    errors.too_large.clone()
                } else {
                    errors.invalid_entry.clone()
                }
            })?,
        _ => {
            return Err(ProjectionError::UnsupportedCompression(
                entry.method.as_u16(),
            ));
        }
    };
    let expected_size = usize::try_from(entry.uncompressed_size).map_err(|_| errors.too_large)?;
    if output.len() != expected_size || crc32(&output) != entry.crc32 {
        return Err(errors.integrity);
    }
    if u64::try_from(local.data().len()).ok() != Some(entry.compressed_size) {
        return Err(errors.integrity);
    }
    Ok(output)
}

fn validate_declared_sizes(
    compressed_size: u64,
    uncompressed_size: u64,
    maximum_bytes: usize,
    too_large: ProjectionError,
    limits: DocxLimits,
) -> Result<(), ProjectionError> {
    let uncompressed_size_as_usize =
        usize::try_from(uncompressed_size).map_err(|_| too_large.clone())?;
    if uncompressed_size_as_usize > maximum_bytes {
        return Err(too_large);
    }
    if uncompressed_size == 0 {
        return Ok(());
    }
    if compressed_size == 0
        || uncompressed_size > compressed_size.saturating_mul(limits.maximum_compression_ratio)
    {
        return Err(ProjectionError::SuspiciousCompressionRatio);
    }
    Ok(())
}
