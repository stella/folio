use miniz_oxide::inflate::decompress_to_vec_with_limit;
use rawzip::{CompressionMethod, ZipArchive, ZipSliceArchive, crc32};

use crate::ProjectionError;
use crate::projection::relationships::main_document_path;

const DOCUMENT_XML_PATH: &[u8] = b"word/document.xml";
const STYLES_XML_PATH: &[u8] = b"word/styles.xml";
const ROOT_RELATIONSHIPS_PATH: &[u8] = b"_rels/.rels";
const MAXIMUM_ROOT_RELATIONSHIPS_BYTES: usize = 1024 * 1024;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DocxLimits {
    pub maximum_archive_bytes: usize,
    pub maximum_document_xml_bytes: usize,
    pub maximum_styles_xml_bytes: usize,
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
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DocumentParts {
    pub document_xml: Vec<u8>,
    pub styles_xml: Option<Vec<u8>>,
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

/// Extracts the main document and its conventional paragraph-style part from one
/// bounded scan of the ZIP directory.
///
/// `styles_xml: None` is not evidence that a package has no effective styles:
/// relationship targets may use another path. Callers therefore retain an
/// explicit unknown state for style-derived structural facts when it is absent.
///
/// # Errors
///
/// Returns [`ProjectionError`] when the package is invalid, unsupported, or
/// exceeds the supplied resource limits.
pub fn extract_document_parts(
    bytes: &[u8],
    limits: DocxLimits,
) -> Result<DocumentParts, ProjectionError> {
    if bytes.len() > limits.maximum_archive_bytes {
        return Err(ProjectionError::ArchiveTooLarge);
    }
    let archive = ZipArchive::from_slice(bytes).map_err(|_| ProjectionError::InvalidArchive)?;
    let package_entries = scan_package_entries(&archive, limits)?;

    let document_path = resolve_document_path(&archive, &package_entries, limits)?;
    let document = unique_entry(
        &package_entries,
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
        &archive,
        document,
        &document_path,
        limits.maximum_document_xml_bytes,
        EntryExtractionErrors::DOCUMENT_XML,
    )?;
    let styles_xml = unique_entry(
        &package_entries,
        STYLES_XML_PATH,
        ProjectionError::DuplicateStylesXml,
    )?
    .map(|styles| {
        validate_selected_entry(
            styles,
            limits.maximum_styles_xml_bytes,
            ProjectionError::StylesXmlTooLarge,
            ProjectionError::InvalidStylesXmlEntry,
            limits,
        )?;
        extract_entry(
            &archive,
            styles,
            STYLES_XML_PATH,
            limits.maximum_styles_xml_bytes,
            EntryExtractionErrors::STYLES_XML,
        )
    })
    .transpose()?;
    Ok(DocumentParts {
        document_xml,
        styles_xml,
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
            .map_err(|_| errors.invalid_entry.clone())?,
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
