use miniz_oxide::inflate::decompress_to_vec_with_limit;
use rawzip::{CompressionMethod, ZipArchive, ZipSliceArchive, crc32};

use crate::ProjectionError;

const DOCUMENT_XML_PATH: &[u8] = b"word/document.xml";
const STYLES_XML_PATH: &[u8] = b"word/styles.xml";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DocxLimits {
    pub maximum_archive_bytes: usize,
    pub maximum_document_xml_bytes: usize,
    pub maximum_styles_xml_bytes: usize,
    pub maximum_entries: u64,
    pub maximum_compression_ratio: u64,
    pub maximum_paragraphs: usize,
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
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DocumentParts {
    pub document_xml: Vec<u8>,
    pub styles_xml: Option<Vec<u8>>,
}

/// Extracts `word/document.xml` from a bounded DOCX package.
///
/// # Errors
///
/// Returns [`ProjectionError`] when the package is invalid, unsupported, or
/// exceeds the supplied resource limits.
pub fn extract_document_xml(bytes: &[u8], limits: DocxLimits) -> Result<Vec<u8>, ProjectionError> {
    extract_document_parts(bytes, limits).map(|parts| parts.document_xml)
}

/// Extracts the main document and its conventional paragraph-style part in one
/// bounded pass over the ZIP directory.
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
    if archive.entries_hint() > limits.maximum_entries {
        return Err(ProjectionError::TooManyArchiveEntries);
    }

    let mut document = None;
    let mut styles = None;
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
        let path = entry.file_path();
        let path = path.as_ref();
        if path != DOCUMENT_XML_PATH && path != STYLES_XML_PATH {
            continue;
        }
        let slot = if path == DOCUMENT_XML_PATH {
            &mut document
        } else {
            &mut styles
        };
        if slot.is_some() {
            return Err(if path == DOCUMENT_XML_PATH {
                ProjectionError::DuplicateDocumentXml
            } else {
                ProjectionError::DuplicateStylesXml
            });
        }
        if entry.flags().is_encrypted() {
            return Err(ProjectionError::EncryptedDocumentXml);
        }
        let method = entry.compression_method();
        if method != CompressionMethod::STORE && method != CompressionMethod::DEFLATE {
            return Err(ProjectionError::UnsupportedCompression(method.as_u16()));
        }
        let compressed_size = entry.compressed_size_hint();
        let uncompressed_size = entry.uncompressed_size_hint();
        let (maximum_bytes, too_large) = if path == DOCUMENT_XML_PATH {
            (
                limits.maximum_document_xml_bytes,
                ProjectionError::DocumentXmlTooLarge,
            )
        } else {
            (
                limits.maximum_styles_xml_bytes,
                ProjectionError::StylesXmlTooLarge,
            )
        };
        validate_declared_sizes(
            compressed_size,
            uncompressed_size,
            maximum_bytes,
            too_large,
            limits,
        )?;
        *slot = Some(XmlEntry {
            wayfinder: entry.wayfinder(),
            method,
            compressed_size,
            uncompressed_size,
            crc32: entry.crc32(),
        });
    }

    let document = document.ok_or(ProjectionError::MissingDocumentXml)?;
    let document_xml = extract_entry(
        &archive,
        document,
        DOCUMENT_XML_PATH,
        limits.maximum_document_xml_bytes,
        ProjectionError::InvalidDocumentXmlEntry,
        ProjectionError::DocumentXmlTooLarge,
        ProjectionError::DocumentXmlIntegrity,
    )?;
    let styles_xml = styles
        .map(|styles| {
            extract_entry(
                &archive,
                styles,
                STYLES_XML_PATH,
                limits.maximum_styles_xml_bytes,
                ProjectionError::InvalidStylesXmlEntry,
                ProjectionError::StylesXmlTooLarge,
                ProjectionError::StylesXmlIntegrity,
            )
        })
        .transpose()?;
    Ok(DocumentParts {
        document_xml,
        styles_xml,
    })
}

fn extract_entry(
    archive: &ZipSliceArchive<&[u8]>,
    entry: XmlEntry,
    expected_path: &[u8],
    maximum_bytes: usize,
    invalid_entry: ProjectionError,
    too_large: ProjectionError,
    integrity: ProjectionError,
) -> Result<Vec<u8>, ProjectionError> {
    let local = archive
        .get_entry(entry.wayfinder)
        .map_err(|_| invalid_entry.clone())?;
    let local_header = local.local_header();
    if local_header.compression_method() != entry.method
        || local_header.file_path().as_ref() != expected_path
        || local_header.flags().is_encrypted()
    {
        return Err(invalid_entry);
    }
    let output = match entry.method {
        CompressionMethod::STORE => local.data().to_vec(),
        CompressionMethod::DEFLATE => decompress_to_vec_with_limit(local.data(), maximum_bytes)
            .map_err(|_| invalid_entry.clone())?,
        _ => {
            return Err(ProjectionError::UnsupportedCompression(
                entry.method.as_u16(),
            ));
        }
    };
    let expected_size = usize::try_from(entry.uncompressed_size).map_err(|_| too_large)?;
    if output.len() != expected_size || crc32(&output) != entry.crc32 {
        return Err(integrity);
    }
    if u64::try_from(local.data().len()).ok() != Some(entry.compressed_size) {
        return Err(integrity);
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
