use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::ProjectionError;

const PACKAGE_RELATIONSHIPS_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/relationships";
const PACKAGE_RELATIONSHIPS_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/package/relationships";
const OFFICE_DOCUMENT_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const OFFICE_DOCUMENT_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
const COMMENTS_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
const COMMENTS_EXTENDED: &str =
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";

#[derive(Debug, Eq, PartialEq)]
pub(super) struct ReviewPartPaths {
    pub comments: Option<Vec<u8>>,
    pub comments_extended: Option<Vec<u8>>,
}

pub(super) fn main_document_path(xml: &[u8]) -> Result<Vec<u8>, ProjectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0_usize;
    let mut root_seen = false;
    let mut target = None;

    loop {
        match reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidPackageRelationships)?
        {
            Event::Start(element) => {
                inspect_element(&reader, &element, depth, &mut root_seen, &mut target)?;
                depth = depth
                    .checked_add(1)
                    .ok_or(ProjectionError::InvalidPackageRelationships)?;
            }
            Event::Empty(element) => {
                inspect_element(&reader, &element, depth, &mut root_seen, &mut target)?;
            }
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or(ProjectionError::InvalidPackageRelationships)?;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !root_seen || depth != 0 {
        return Err(ProjectionError::InvalidPackageRelationships);
    }
    target.ok_or(ProjectionError::MissingDocumentXml)
}

pub(super) fn document_relationships_path(
    document_path: &[u8],
) -> Result<Vec<u8>, ProjectionError> {
    let path = std::str::from_utf8(document_path)
        .map_err(|_| ProjectionError::InvalidPackageRelationships)?;
    let (parent, file_name) = path.rsplit_once('/').unwrap_or(("", path));
    if file_name.is_empty() {
        return Err(ProjectionError::InvalidPackageRelationships);
    }
    let prefix = if parent.is_empty() {
        String::new()
    } else {
        format!("{parent}/")
    };
    Ok(format!("{prefix}_rels/{file_name}.rels").into_bytes())
}

pub(super) fn review_part_paths(
    xml: &[u8],
    document_path: &[u8],
) -> Result<ReviewPartPaths, ProjectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0_usize;
    let mut root_seen = false;
    let mut paths = ReviewPartPaths {
        comments: None,
        comments_extended: None,
    };

    loop {
        match reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidPackageRelationships)?
        {
            Event::Start(element) => {
                inspect_review_element(
                    &reader,
                    &element,
                    depth,
                    &mut root_seen,
                    &mut paths,
                    document_path,
                )?;
                depth = depth
                    .checked_add(1)
                    .ok_or(ProjectionError::InvalidPackageRelationships)?;
            }
            Event::Empty(element) => inspect_review_element(
                &reader,
                &element,
                depth,
                &mut root_seen,
                &mut paths,
                document_path,
            )?,
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or(ProjectionError::InvalidPackageRelationships)?;
            }
            Event::DocType(_) => return Err(ProjectionError::InvalidPackageRelationships),
            Event::Eof => break,
            _ => {}
        }
    }
    if !root_seen || depth != 0 {
        return Err(ProjectionError::InvalidPackageRelationships);
    }
    Ok(paths)
}

fn inspect_review_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    depth: usize,
    root_seen: &mut bool,
    paths: &mut ReviewPartPaths,
    document_path: &[u8],
) -> Result<(), ProjectionError> {
    let (namespace, local_name) = reader.resolver().resolve_element(element.name());
    if depth == 0 {
        if *root_seen
            || local_name.as_ref() != b"Relationships"
            || !is_relationships_namespace(&namespace)
        {
            return Err(ProjectionError::InvalidPackageRelationships);
        }
        *root_seen = true;
        return Ok(());
    }
    if depth != 1
        || local_name.as_ref() != b"Relationship"
        || !is_relationships_namespace(&namespace)
    {
        return Ok(());
    }

    let relationship_type = unqualified_attribute(reader, element, b"Type")?;
    let selected = match relationship_type.as_deref() {
        Some(COMMENTS_TRANSITIONAL | COMMENTS_STRICT) => &mut paths.comments,
        Some(COMMENTS_EXTENDED) => &mut paths.comments_extended,
        _ => return Ok(()),
    };
    if selected.is_some() {
        return Err(ProjectionError::InvalidPackageRelationships);
    }
    match unqualified_attribute(reader, element, b"TargetMode")?.as_deref() {
        None | Some("Internal") => {}
        Some(_) => return Err(ProjectionError::InvalidPackageRelationships),
    }
    let target = unqualified_attribute(reader, element, b"Target")?
        .ok_or(ProjectionError::InvalidPackageRelationships)?;
    *selected = Some(normalize_part_target(document_path, &target)?);
    Ok(())
}

fn inspect_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    depth: usize,
    root_seen: &mut bool,
    target: &mut Option<Vec<u8>>,
) -> Result<(), ProjectionError> {
    let (namespace, local_name) = reader.resolver().resolve_element(element.name());
    if depth == 0 {
        if *root_seen
            || local_name.as_ref() != b"Relationships"
            || !is_relationships_namespace(&namespace)
        {
            return Err(ProjectionError::InvalidPackageRelationships);
        }
        *root_seen = true;
        return Ok(());
    }
    if depth != 1
        || local_name.as_ref() != b"Relationship"
        || !is_relationships_namespace(&namespace)
    {
        return Ok(());
    }

    let relationship_type = unqualified_attribute(reader, element, b"Type")?;
    if !matches!(
        relationship_type.as_deref(),
        Some(OFFICE_DOCUMENT_TRANSITIONAL | OFFICE_DOCUMENT_STRICT)
    ) {
        return Ok(());
    }
    if target.is_some() {
        return Err(ProjectionError::DuplicateDocumentXml);
    }
    match unqualified_attribute(reader, element, b"TargetMode")?.as_deref() {
        None | Some("Internal") => {}
        Some(_) => return Err(ProjectionError::InvalidPackageRelationships),
    }
    let value = unqualified_attribute(reader, element, b"Target")?
        .ok_or(ProjectionError::InvalidPackageRelationships)?;
    *target = Some(normalize_root_target(&value)?);
    Ok(())
}

fn is_relationships_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(value)
            if matches!(
                value.as_ref(),
                PACKAGE_RELATIONSHIPS_TRANSITIONAL | PACKAGE_RELATIONSHIPS_STRICT
            )
    )
}

fn unqualified_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, ProjectionError> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ProjectionError::InvalidPackageRelationships)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if namespace == ResolveResult::Unbound && local_name.as_ref() == name {
            if value.is_some() {
                return Err(ProjectionError::InvalidPackageRelationships);
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|_| ProjectionError::InvalidPackageRelationships)?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn normalize_root_target(target: &str) -> Result<Vec<u8>, ProjectionError> {
    if target.is_empty()
        || target.contains(['\\', '?', '#', '\0'])
        || target.contains(':')
        || !target.is_ascii()
    {
        return Err(ProjectionError::InvalidPackageRelationships);
    }
    let mut segments = Vec::new();
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                if segments.pop().is_none() {
                    return Err(ProjectionError::InvalidPackageRelationships);
                }
            }
            _ => segments.push(segment),
        }
    }
    if segments.is_empty() {
        return Err(ProjectionError::InvalidPackageRelationships);
    }
    Ok(segments.join("/").into_bytes())
}

fn normalize_part_target(source_path: &[u8], target: &str) -> Result<Vec<u8>, ProjectionError> {
    let source = std::str::from_utf8(source_path)
        .map_err(|_| ProjectionError::InvalidPackageRelationships)?;
    if target.starts_with('/') {
        return normalize_root_target(target);
    }
    let parent = source.rsplit_once('/').map_or("", |(parent, _)| parent);
    let combined = if parent.is_empty() {
        target.to_owned()
    } else {
        format!("{parent}/{target}")
    };
    normalize_root_target(&combined)
}

#[cfg(test)]
mod tests {
    use super::{document_relationships_path, review_part_paths};

    #[test]
    fn derives_relationship_and_relative_part_paths_at_any_package_depth() {
        assert_eq!(
            document_relationships_path(b"document.xml"),
            Ok(b"_rels/document.xml.rels".to_vec())
        );
        assert_eq!(
            document_relationships_path(b"custom/main.xml"),
            Ok(b"custom/_rels/main.xml.rels".to_vec())
        );
        let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../review/comments.xml"/></Relationships>"#;
        assert_eq!(
            review_part_paths(relationships, b"custom/main.xml").map(|paths| paths.comments),
            Ok(Some(b"review/comments.xml".to_vec()))
        );
    }

    #[test]
    fn rejects_external_duplicate_and_package_escaping_review_targets() {
        for relationship in [
            r#"<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml" TargetMode="External"/>"#,
            r#"<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/comments" Target="other.xml"/>"#,
            r#"<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../../comments.xml"/>"#,
        ] {
            let xml = format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationship}</Relationships>"#
            );
            assert!(review_part_paths(xml.as_bytes(), b"custom/main.xml").is_err());
        }
    }
}
