use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use crate::ProjectionError;
use crate::projection::namespaces::OoxmlNamespace;
use crate::projection::structure::{
    ParagraphIndentation, ParagraphProperties, StyleDefinition, StyleSheet,
};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum PropertiesOwner {
    DocumentDefaults,
    ParagraphStyle,
    Other,
}

#[derive(Debug)]
enum Frame {
    Other,
    Styles,
    DocumentDefaults,
    ParagraphProperties(PropertiesOwner),
    NumberingProperties(PropertiesOwner),
    Style,
}

struct StyleBuilder {
    id: String,
    based_on: Option<String>,
    properties: ParagraphProperties,
    is_default: bool,
    is_paragraph: bool,
}

pub(super) fn parse_styles(xml: &[u8]) -> Result<StyleSheet, ProjectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut frames = Vec::new();
    let mut sheet = StyleSheet::default();
    let mut current_style: Option<StyleBuilder> = None;
    let mut root_seen = false;

    loop {
        match reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidStylesXml)?
        {
            Event::Start(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                start(
                    &reader,
                    OoxmlNamespace::from_resolved(&namespace),
                    &element,
                    &mut frames,
                    &mut sheet,
                    &mut current_style,
                    &mut root_seen,
                )?;
            }
            Event::Empty(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                start(
                    &reader,
                    OoxmlNamespace::from_resolved(&namespace),
                    &element,
                    &mut frames,
                    &mut sheet,
                    &mut current_style,
                    &mut root_seen,
                )?;
                end(&mut frames, &mut sheet, &mut current_style)?;
            }
            Event::End(_) => end(&mut frames, &mut sheet, &mut current_style)?,
            Event::Eof => break,
            _ => {}
        }
    }
    if !root_seen || !frames.is_empty() || current_style.is_some() {
        return Err(ProjectionError::InvalidStylesXml);
    }
    Ok(sheet)
}

fn start(
    reader: &NsReader<&[u8]>,
    namespace: OoxmlNamespace,
    element: &BytesStart<'_>,
    frames: &mut Vec<Frame>,
    sheet: &mut StyleSheet,
    current_style: &mut Option<StyleBuilder>,
    root_seen: &mut bool,
) -> Result<(), ProjectionError> {
    let local_name = element.local_name();
    let name = if namespace == OoxmlNamespace::Wordprocessing {
        local_name.as_ref()
    } else {
        b""
    };
    let frame = match name {
        b"styles" if frames.is_empty() && !*root_seen => {
            *root_seen = true;
            Frame::Styles
        }
        b"docDefaults" if matches!(frames.last(), Some(Frame::Styles)) => Frame::DocumentDefaults,
        b"style" if matches!(frames.last(), Some(Frame::Styles)) => {
            if current_style.is_some() {
                return Err(ProjectionError::InvalidStylesXml);
            }
            let id =
                attribute(reader, element, b"styleId")?.ok_or(ProjectionError::InvalidStylesXml)?;
            let style_type = attribute(reader, element, b"type")?;
            let is_default = on_off_attribute(reader, element, b"default")?;
            *current_style = Some(StyleBuilder {
                id,
                based_on: None,
                properties: ParagraphProperties::default(),
                is_default,
                is_paragraph: style_type.as_deref() == Some("paragraph"),
            });
            Frame::Style
        }
        b"basedOn" if matches!(frames.last(), Some(Frame::Style)) => {
            if let Some(style) = current_style.as_mut() {
                style.based_on = attribute(reader, element, b"val")?;
            }
            Frame::Other
        }
        b"pPr" => {
            let owner = if current_style
                .as_ref()
                .is_some_and(|style| style.is_paragraph)
                && frames
                    .iter()
                    .rev()
                    .any(|frame| matches!(frame, Frame::Style))
            {
                PropertiesOwner::ParagraphStyle
            } else if frames
                .iter()
                .rev()
                .any(|frame| matches!(frame, Frame::DocumentDefaults))
            {
                PropertiesOwner::DocumentDefaults
            } else {
                PropertiesOwner::Other
            };
            Frame::ParagraphProperties(owner)
        }
        b"numPr" => {
            let owner = paragraph_properties_owner(frames).unwrap_or(PropertiesOwner::Other);
            if let Some(properties) = properties_mut(owner, sheet, current_style) {
                properties.numbering.present = true;
            }
            Frame::NumberingProperties(owner)
        }
        b"ind" => {
            if let Some(owner) = paragraph_properties_owner(frames)
                && let Some(properties) = properties_mut(owner, sheet, current_style)
            {
                properties.indentation = parse_indentation(reader, element)?;
            }
            Frame::Other
        }
        b"numId" => {
            if let Some(owner) = numbering_properties_owner(frames)
                && let Some(properties) = properties_mut(owner, sheet, current_style)
            {
                properties.numbering.num_id = Some(parse_u32_attribute(reader, element)?);
            }
            Frame::Other
        }
        b"ilvl" => {
            if let Some(owner) = numbering_properties_owner(frames)
                && let Some(properties) = properties_mut(owner, sheet, current_style)
            {
                properties.numbering.level = Some(parse_level_attribute(reader, element)?);
            }
            Frame::Other
        }
        _ => Frame::Other,
    };
    frames.push(frame);
    Ok(())
}

fn end(
    frames: &mut Vec<Frame>,
    sheet: &mut StyleSheet,
    current_style: &mut Option<StyleBuilder>,
) -> Result<(), ProjectionError> {
    let frame = frames.pop().ok_or(ProjectionError::InvalidStylesXml)?;
    if matches!(frame, Frame::Style) {
        let style = current_style
            .take()
            .ok_or(ProjectionError::InvalidStylesXml)?;
        if style.is_paragraph {
            if style.is_default && sheet.default_style_id.replace(style.id.clone()).is_some() {
                return Err(ProjectionError::InvalidStylesXml);
            }
            if sheet
                .styles
                .insert(
                    style.id,
                    StyleDefinition {
                        based_on: style.based_on,
                        properties: style.properties,
                    },
                )
                .is_some()
            {
                return Err(ProjectionError::InvalidStylesXml);
            }
        }
    }
    Ok(())
}

fn paragraph_properties_owner(frames: &[Frame]) -> Option<PropertiesOwner> {
    frames.iter().rev().find_map(|frame| match frame {
        Frame::ParagraphProperties(owner) => Some(*owner),
        _ => None,
    })
}

fn numbering_properties_owner(frames: &[Frame]) -> Option<PropertiesOwner> {
    frames.iter().rev().find_map(|frame| match frame {
        Frame::NumberingProperties(owner) => Some(*owner),
        _ => None,
    })
}

fn properties_mut<'a>(
    owner: PropertiesOwner,
    sheet: &'a mut StyleSheet,
    current_style: &'a mut Option<StyleBuilder>,
) -> Option<&'a mut ParagraphProperties> {
    match owner {
        PropertiesOwner::DocumentDefaults => Some(&mut sheet.document_defaults),
        PropertiesOwner::ParagraphStyle => {
            current_style.as_mut().map(|style| &mut style.properties)
        }
        PropertiesOwner::Other => None,
    }
}

pub(super) fn parse_indentation(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ParagraphIndentation, ProjectionError> {
    Ok(ParagraphIndentation {
        first_line_twips: signed_attribute(reader, element, b"firstLine")?,
        hanging_twips: signed_attribute(reader, element, b"hanging")?,
        left_twips: signed_attribute(reader, element, b"left")?,
        right_twips: signed_attribute(reader, element, b"right")?,
        start_twips: signed_attribute(reader, element, b"start")?,
        end_twips: signed_attribute(reader, element, b"end")?,
        first_line_chars_hundredths: signed_attribute(reader, element, b"firstLineChars")?,
        hanging_chars_hundredths: signed_attribute(reader, element, b"hangingChars")?,
        left_chars_hundredths: signed_attribute(reader, element, b"leftChars")?,
        right_chars_hundredths: signed_attribute(reader, element, b"rightChars")?,
        start_chars_hundredths: signed_attribute(reader, element, b"startChars")?,
        end_chars_hundredths: signed_attribute(reader, element, b"endChars")?,
    })
}

pub(super) fn parse_u32_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<u32, ProjectionError> {
    attribute(reader, element, b"val")?
        .ok_or(ProjectionError::InvalidDocumentXml)?
        .parse()
        .map_err(|_| ProjectionError::InvalidDocumentXml)
}

pub(super) fn parse_level_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<u8, ProjectionError> {
    let level = parse_u32_attribute(reader, element)?;
    u8::try_from(level).map_err(|_| ProjectionError::InvalidDocumentXml)
}

fn signed_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<i32>, ProjectionError> {
    attribute(reader, element, name)?
        .map(|value| {
            value
                .parse()
                .map_err(|_| ProjectionError::InvalidDocumentXml)
        })
        .transpose()
}

fn on_off_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<bool, ProjectionError> {
    Ok(attribute(reader, element, name)?
        .is_some_and(|value| !matches!(value.to_ascii_lowercase().as_str(), "0" | "false" | "off")))
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, ProjectionError> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ProjectionError::InvalidStylesXml)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if OoxmlNamespace::from_resolved(&namespace) == OoxmlNamespace::Wordprocessing
            && local_name.as_ref() == name
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|_| ProjectionError::InvalidStylesXml);
        }
    }
    Ok(None)
}
