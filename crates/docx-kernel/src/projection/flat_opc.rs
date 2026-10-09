use std::ops::Range;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{PrefixDeclaration, ResolveResult};
use quick_xml::reader::{NsReader, Reader};

use super::relationships::{
    document_relationship_paths, document_relationships_path, main_document_path,
};
use super::{DocumentParts, DocxLimits, ProjectionError};

const PACKAGE_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/2006/xmlPackage";
const ROOT_RELATIONSHIPS_PATH: &str = "/_rels/.rels";
const MAXIMUM_RELATIONSHIPS_BYTES: usize = 1024 * 1024;

#[derive(Clone, Copy, Eq, PartialEq)]
pub(super) enum XmlInputKind {
    Document,
    Package,
}

pub(super) fn input_kind(xml: &[u8]) -> Result<XmlInputKind, ProjectionError> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidDocumentXml)?
        {
            Event::Start(root) | Event::Empty(root) => {
                let (namespace, local) = reader.resolver().resolve_element(root.name());
                if local.as_ref() == b"document"
                    && super::namespaces::OoxmlNamespace::from_resolved(&namespace)
                        == super::namespaces::OoxmlNamespace::Wordprocessing
                {
                    return Ok(XmlInputKind::Document);
                }
                if local.as_ref() == b"package"
                    && matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == PACKAGE_NAMESPACE)
                {
                    return Ok(XmlInputKind::Package);
                }
                return Err(ProjectionError::InvalidDocumentXml);
            }
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {}
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {}
            _ => return Err(ProjectionError::InvalidDocumentXml),
        }
    }
}

struct XmlPart {
    range: Range<usize>,
    namespaces: Vec<(Vec<u8>, Vec<u8>)>,
}

pub(super) fn extract_parts(
    xml: &[u8],
    limits: DocxLimits,
) -> Result<DocumentParts, ProjectionError> {
    if xml.len() > limits.maximum_archive_bytes {
        return Err(ProjectionError::ArchiveTooLarge);
    }
    let parts = index_parts(xml, limits)?;
    let relationships = extract(
        xml,
        find_part(&parts, ROOT_RELATIONSHIPS_PATH).ok_or(ProjectionError::MissingDocumentXml)?,
        MAXIMUM_RELATIONSHIPS_BYTES,
        ProjectionError::PackageRelationshipsTooLarge,
    )?;
    let document_path = main_document_path(&relationships)?;
    let document_name = part_name(&document_path)?;
    let document_xml = extract(
        xml,
        find_part(&parts, &document_name).ok_or(ProjectionError::MissingDocumentXml)?,
        limits.maximum_document_xml_bytes,
        ProjectionError::DocumentXmlTooLarge,
    )?;
    if input_kind(&document_xml)? != XmlInputKind::Document {
        return Err(ProjectionError::InvalidDocumentXml);
    }
    let relationship_name = part_name(&document_relationships_path(&document_path)?)?;
    let related = find_part(&parts, &relationship_name)
        .map(|part| {
            extract(
                xml,
                part,
                MAXIMUM_RELATIONSHIPS_BYTES,
                ProjectionError::PackageRelationshipsTooLarge,
            )
        })
        .transpose()?
        .map(|bytes| document_relationship_paths(&bytes, &document_path))
        .transpose()?
        .unwrap_or_default();
    let styles_xml = extract_optional(
        xml,
        &parts,
        related.styles.as_deref(),
        limits.maximum_styles_xml_bytes,
        ProjectionError::StylesXmlTooLarge,
    )?;
    let numbering_xml = extract_optional(
        xml,
        &parts,
        related.numbering.as_deref(),
        limits.maximum_numbering_xml_bytes,
        ProjectionError::NumberingXmlTooLarge,
    )?;
    Ok(DocumentParts {
        document_xml,
        styles_xml,
        numbering_xml,
    })
}

fn find_part<'a>(parts: &'a [(String, XmlPart)], name: &str) -> Option<&'a XmlPart> {
    parts
        .iter()
        .find(|(path, _)| path == name)
        .map(|(_, part)| part)
}

fn part_name(path: &[u8]) -> Result<String, ProjectionError> {
    Ok(format!(
        "/{}",
        std::str::from_utf8(path).map_err(|_| ProjectionError::InvalidFlatOpcPackage)?
    ))
}

fn extract_optional(
    xml: &[u8],
    parts: &[(String, XmlPart)],
    path: Option<&[u8]>,
    maximum: usize,
    error: ProjectionError,
) -> Result<Option<Vec<u8>>, ProjectionError> {
    let Some(path) = path else {
        return Ok(None);
    };
    let name = part_name(path)?;
    find_part(parts, &name)
        .map(|part| extract(xml, part, maximum, error))
        .transpose()
}

// Index package topology only: no body search, payload projection or binary decoding.
#[derive(Default)]
struct PartIndex {
    parts: Vec<(String, XmlPart)>,
    names: std::collections::HashSet<String>,
    depth: usize,
    root_seen: bool,
    current_name: Option<String>,
    current_xml: Option<XmlPart>,
    payload_seen: bool,
    index_bytes: usize,
}

impl PartIndex {
    fn start(
        &mut self,
        reader: &mut NsReader<&[u8]>,
        element: &BytesStart<'_>,
        limits: DocxLimits,
    ) -> Result<(), ProjectionError> {
        let (namespace, local) = reader.resolver().resolve_element(element.name());
        if !matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == PACKAGE_NAMESPACE) {
            return Err(ProjectionError::InvalidFlatOpcPackage);
        }
        match (self.depth, local.as_ref()) {
            (0, b"package") if !self.root_seen => {
                self.root_seen = true;
                self.depth = 1;
            }
            (1, b"part") => {
                let name = package_attribute(reader, element, b"name")?
                    .ok_or(ProjectionError::InvalidFlatOpcPackage)?;
                if !name.starts_with('/')
                    || name.contains('\\')
                    || name
                        .split('/')
                        .skip(1)
                        .any(|segment| segment.is_empty() || segment == "." || segment == "..")
                    || !self.names.insert(name.clone())
                {
                    return Err(ProjectionError::InvalidFlatOpcPackage);
                }
                if u64::try_from(self.names.len())
                    .map_err(|_| ProjectionError::TooManyArchiveEntries)?
                    > limits.maximum_entries
                {
                    return Err(ProjectionError::TooManyArchiveEntries);
                }
                self.index_bytes = self
                    .index_bytes
                    .checked_add(name.len())
                    .ok_or(ProjectionError::ArchiveTooLarge)?;
                self.current_name = Some(name);
                self.payload_seen = false;
                self.depth = 2;
            }
            (2, b"xmlData" | b"binaryData") if !self.payload_seen => {
                self.payload_seen = true;
                let namespaces = reader
                    .resolver()
                    .bindings()
                    .map(|(prefix, namespace)| {
                        let name = match prefix {
                            PrefixDeclaration::Default => b"xmlns".to_vec(),
                            PrefixDeclaration::Named(prefix) => {
                                [b"xmlns:".as_slice(), prefix].concat()
                            }
                        };
                        (name, namespace.as_ref().to_vec())
                    })
                    .collect::<Vec<_>>();
                for (name, value) in &namespaces {
                    self.index_bytes = self
                        .index_bytes
                        .checked_add(name.len())
                        .and_then(|size| size.checked_add(value.len()))
                        .ok_or(ProjectionError::ArchiveTooLarge)?;
                }
                if self.index_bytes > limits.maximum_archive_bytes {
                    return Err(ProjectionError::ArchiveTooLarge);
                }
                let is_xml = local.as_ref() == b"xmlData";
                let span = reader
                    .read_to_end(element.name())
                    .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?;
                if is_xml {
                    self.current_xml = Some(XmlPart {
                        range: usize::try_from(span.start)
                            .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?
                            ..usize::try_from(span.end)
                                .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?,
                        namespaces,
                    });
                }
            }
            _ => return Err(ProjectionError::InvalidFlatOpcPackage),
        }
        Ok(())
    }

    fn empty(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let (namespace, local) = reader.resolver().resolve_element(element.name());
        if !matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == PACKAGE_NAMESPACE) {
            return Err(ProjectionError::InvalidFlatOpcPackage);
        }
        match (self.depth, local.as_ref()) {
            (0, b"package") if !self.root_seen => self.root_seen = true,
            (2, b"xmlData" | b"binaryData") if !self.payload_seen => {
                self.payload_seen = true;
                if local.as_ref() == b"xmlData" {
                    let position = usize::try_from(reader.buffer_position())
                        .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?;
                    self.current_xml = Some(XmlPart {
                        range: position..position,
                        namespaces: Vec::new(),
                    });
                }
            }
            _ => return Err(ProjectionError::InvalidFlatOpcPackage),
        }
        Ok(())
    }

    fn end(&mut self) -> Result<(), ProjectionError> {
        if self.depth == 2 {
            let name = self
                .current_name
                .take()
                .ok_or(ProjectionError::InvalidFlatOpcPackage)?;
            if !self.payload_seen {
                return Err(ProjectionError::InvalidFlatOpcPackage);
            }
            if let Some(part) = self.current_xml.take() {
                self.parts.push((name, part));
            }
        }
        self.depth = self
            .depth
            .checked_sub(1)
            .ok_or(ProjectionError::InvalidFlatOpcPackage)?;
        Ok(())
    }
}

fn index_parts(xml: &[u8], limits: DocxLimits) -> Result<Vec<(String, XmlPart)>, ProjectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut state = PartIndex::default();
    loop {
        match reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?
        {
            Event::Start(element) => state.start(&mut reader, &element, limits)?,
            Event::Empty(element) => state.empty(&reader, &element)?,
            Event::End(_) => state.end()?,
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {}
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {}
            Event::Eof if state.root_seen && state.depth == 0 => return Ok(state.parts),
            _ => return Err(ProjectionError::InvalidFlatOpcPackage),
        }
    }
}

fn package_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    key: &[u8],
) -> Result<Option<String>, ProjectionError> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ProjectionError::InvalidFlatOpcPackage)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == key
            && matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == PACKAGE_NAMESPACE)
        {
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn extract(
    xml: &[u8],
    part: &XmlPart,
    maximum: usize,
    too_large: ProjectionError,
) -> Result<Vec<u8>, ProjectionError> {
    let bytes = xml
        .get(part.range.clone())
        .ok_or(ProjectionError::InvalidFlatOpcPackage)?;
    if bytes.len() > maximum {
        return Err(too_large);
    }
    if bytes.is_empty() {
        return Ok(Vec::new());
    }
    let mut reader = Reader::from_reader(bytes);
    loop {
        let event = reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?;
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(mut root) | Event::Empty(mut root) => {
                let declared = root
                    .attributes()
                    .map(|attribute| attribute.map(|value| value.key.as_ref().to_vec()))
                    .collect::<Result<Vec<_>, _>>()
                    .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?;
                for (name, namespace) in &part.namespaces {
                    if !declared.contains(name) {
                        root.push_attribute((name.as_slice(), namespace.as_slice()));
                    }
                }
                let offset = usize::try_from(reader.buffer_position())
                    .map_err(|_| ProjectionError::InvalidFlatOpcPackage)?;
                let tail = bytes
                    .get(offset..)
                    .ok_or(ProjectionError::InvalidFlatOpcPackage)?;
                let closing = if empty {
                    b"/>".as_slice()
                } else {
                    b">".as_slice()
                };
                let size = root
                    .as_ref()
                    .len()
                    .checked_add(if empty { 3 } else { 2 })
                    .and_then(|size| size.checked_add(tail.len()))
                    .filter(|size| *size <= maximum)
                    .ok_or(too_large)?;
                let mut output = Vec::with_capacity(size);
                output.push(b'<');
                output.extend_from_slice(root.as_ref());
                output.extend_from_slice(closing);
                output.extend_from_slice(tail);
                return Ok(output);
            }
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {}
            Event::Comment(_) | Event::PI(_) => {}
            _ => return Err(ProjectionError::InvalidFlatOpcPackage),
        }
    }
}
