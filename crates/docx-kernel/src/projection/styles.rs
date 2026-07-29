use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use crate::ProjectionError;
use crate::projection::compatibility::{CompatibilityAction, MarkupCompatibility};
use crate::projection::namespaces::OoxmlNamespace;
use crate::projection::structure::{
    ParagraphIndentation, ParagraphProperties, StyleDefinition, StyleKind, StyleSheet,
    TextProperties,
};

const MAXIMUM_OUTLINE_LEVEL: u8 = 9;
const MAXIMUM_WORD_STYLE_ID_CHARACTERS: usize = 253;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum PropertiesOwner {
    DocumentDefaults,
    Style,
    Other,
}

#[derive(Debug)]
enum Frame {
    Other,
    Styles,
    DocumentDefaults,
    ParagraphPropertiesDefault,
    RunPropertiesDefault,
    ParagraphProperties(PropertiesOwner),
    RunProperties(PropertiesOwner),
    NumberingProperties(PropertiesOwner),
    Style,
}

struct StyleBuilder {
    id: String,
    based_on: Option<String>,
    properties: ParagraphProperties,
    text_properties: TextProperties,
    is_default: bool,
    kind: Option<StyleKind>,
}

struct StyleParserState {
    frames: Vec<Frame>,
    sheet: StyleSheet,
    current_style: Option<StyleBuilder>,
    root_seen: bool,
    style_count: usize,
    maximum_styles: usize,
}

impl StyleParserState {
    fn new(maximum_styles: usize) -> Self {
        Self {
            frames: Vec::new(),
            sheet: StyleSheet::default(),
            current_style: None,
            root_seen: false,
            style_count: 0,
            maximum_styles,
        }
    }
}

pub(super) fn parse_styles(
    xml: &[u8],
    maximum_styles: usize,
) -> Result<StyleSheet, ProjectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut state = StyleParserState::new(maximum_styles);
    let mut compatibility = MarkupCompatibility::for_styles();

    loop {
        match reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidStylesXml)?
        {
            Event::Start(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                let namespace = OoxmlNamespace::from_resolved(&namespace);
                if compatibility.start(&reader, namespace, &element)?
                    == CompatibilityAction::Process
                {
                    start(&reader, namespace, &element, &mut state)?;
                }
            }
            Event::Empty(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                let namespace = OoxmlNamespace::from_resolved(&namespace);
                if compatibility.empty(&reader, namespace, &element)?
                    == CompatibilityAction::Process
                {
                    start(&reader, namespace, &element, &mut state)?;
                    end(&mut state)?;
                }
            }
            Event::End(element) => {
                let (namespace, local_name) = reader.resolver().resolve_element(element.name());
                if compatibility.end(
                    OoxmlNamespace::from_resolved(&namespace),
                    local_name.as_ref(),
                )? == CompatibilityAction::Process
                {
                    end(&mut state)?;
                }
            }
            Event::DocType(_) => return Err(ProjectionError::InvalidStylesXml),
            Event::Eof => break,
            _ => {}
        }
    }
    if !state.root_seen
        || !state.frames.is_empty()
        || state.current_style.is_some()
        || !compatibility.is_complete()
    {
        return Err(ProjectionError::InvalidStylesXml);
    }
    state.sheet.prepare_styles();
    Ok(state.sheet)
}

fn start(
    reader: &NsReader<&[u8]>,
    namespace: OoxmlNamespace,
    element: &BytesStart<'_>,
    state: &mut StyleParserState,
) -> Result<(), ProjectionError> {
    let local_name = element.local_name();
    let name = if namespace == OoxmlNamespace::Wordprocessing {
        local_name.as_ref()
    } else {
        b""
    };
    if parse_text_property(
        name,
        reader,
        element,
        &state.frames,
        &mut state.sheet,
        &mut state.current_style,
    )? {
        state.frames.push(Frame::Other);
        return Ok(());
    }
    let frame = match name {
        b"styles" if state.frames.is_empty() && !state.root_seen => {
            state.root_seen = true;
            Frame::Styles
        }
        b"docDefaults" if matches!(state.frames.last(), Some(Frame::Styles)) => {
            Frame::DocumentDefaults
        }
        b"style" if matches!(state.frames.last(), Some(Frame::Styles)) => {
            start_style_definition(reader, element, state)?
        }
        b"basedOn" if matches!(state.frames.last(), Some(Frame::Style)) => {
            if let Some(style) = state.current_style.as_mut() {
                style.based_on = word_style_id(attribute(reader, element, b"val")?);
            }
            Frame::Other
        }
        b"pPrDefault" if matches!(state.frames.last(), Some(Frame::DocumentDefaults)) => {
            Frame::ParagraphPropertiesDefault
        }
        b"rPrDefault" if matches!(state.frames.last(), Some(Frame::DocumentDefaults)) => {
            Frame::RunPropertiesDefault
        }
        b"pPr" if matches!(state.frames.last(), Some(Frame::ParagraphPropertiesDefault)) => {
            Frame::ParagraphProperties(PropertiesOwner::DocumentDefaults)
        }
        b"pPr"
            if matches!(state.frames.last(), Some(Frame::Style))
                && state
                    .current_style
                    .as_ref()
                    .is_some_and(|style| style.kind == Some(StyleKind::Paragraph)) =>
        {
            Frame::ParagraphProperties(PropertiesOwner::Style)
        }
        b"rPr" if matches!(state.frames.last(), Some(Frame::RunPropertiesDefault)) => {
            Frame::RunProperties(PropertiesOwner::DocumentDefaults)
        }
        b"rPr"
            if matches!(state.frames.last(), Some(Frame::Style))
                && state.current_style.is_some() =>
        {
            Frame::RunProperties(PropertiesOwner::Style)
        }
        _ => paragraph_property_frame(name, reader, element, state)?,
    };
    state.frames.push(frame);
    Ok(())
}

fn start_style_definition(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    state: &mut StyleParserState,
) -> Result<Frame, ProjectionError> {
    if state.current_style.is_some() {
        return Err(ProjectionError::InvalidStylesXml);
    }
    state.style_count = state
        .style_count
        .checked_add(1)
        .ok_or(ProjectionError::InvalidStylesXml)?;
    if state.style_count > state.maximum_styles {
        return Err(ProjectionError::InvalidStylesXml);
    }
    let id = attribute(reader, element, b"styleId")?.ok_or(ProjectionError::InvalidStylesXml)?;
    let id_supported = id.chars().count() <= MAXIMUM_WORD_STYLE_ID_CHARACTERS;
    let style_type = attribute(reader, element, b"type")?;
    let is_default = on_off_attribute(reader, element, b"default")?;
    state.current_style = Some(StyleBuilder {
        id,
        based_on: None,
        properties: ParagraphProperties::default(),
        text_properties: TextProperties::default(),
        is_default,
        kind: match (id_supported, style_type.as_deref()) {
            (true, Some("character")) => Some(StyleKind::Character),
            (true, Some("paragraph")) => Some(StyleKind::Paragraph),
            _ => None,
        },
    });
    Ok(Frame::Style)
}

pub(super) fn word_style_id(style_id: Option<String>) -> Option<String> {
    style_id.filter(|value| value.chars().count() <= MAXIMUM_WORD_STYLE_ID_CHARACTERS)
}

fn paragraph_property_frame(
    name: &[u8],
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    state: &mut StyleParserState,
) -> Result<Frame, ProjectionError> {
    let frame = match name {
        b"numPr" => {
            let owner = paragraph_properties_owner(&state.frames).unwrap_or(PropertiesOwner::Other);
            if let Some(properties) =
                properties_mut(owner, &mut state.sheet, &mut state.current_style)
            {
                properties.numbering.present = true;
            }
            Frame::NumberingProperties(owner)
        }
        b"ind" => {
            if let Some(owner) = paragraph_properties_owner(&state.frames)
                && let Some(properties) =
                    properties_mut(owner, &mut state.sheet, &mut state.current_style)
            {
                properties.indentation = parse_indentation(reader, element)?;
            }
            Frame::Other
        }
        b"outlineLvl" => {
            if let Some(owner) = paragraph_properties_owner(&state.frames)
                && let Some(properties) =
                    properties_mut(owner, &mut state.sheet, &mut state.current_style)
            {
                properties.outline_level = Some(parse_outline_level_attribute(reader, element)?);
            }
            Frame::Other
        }
        b"numId" => {
            if let Some(owner) = numbering_properties_owner(&state.frames)
                && let Some(properties) =
                    properties_mut(owner, &mut state.sheet, &mut state.current_style)
            {
                properties.numbering.num_id = Some(parse_u32_attribute(reader, element)?);
            }
            Frame::Other
        }
        b"ilvl" => {
            if let Some(owner) = numbering_properties_owner(&state.frames)
                && let Some(properties) =
                    properties_mut(owner, &mut state.sheet, &mut state.current_style)
            {
                properties.numbering.level = Some(parse_level_attribute(reader, element)?);
            }
            Frame::Other
        }
        _ => Frame::Other,
    };
    Ok(frame)
}

fn end(state: &mut StyleParserState) -> Result<(), ProjectionError> {
    let frame = state
        .frames
        .pop()
        .ok_or(ProjectionError::InvalidStylesXml)?;
    if matches!(frame, Frame::Style) {
        let style = state
            .current_style
            .take()
            .ok_or(ProjectionError::InvalidStylesXml)?;
        if let Some(kind) = style.kind {
            let ignored_by_word = matches!(
                style.id.as_str(),
                "NoList" | "DefaultParagraphFont" | "TableNormal"
            );
            if kind == StyleKind::Paragraph
                && style.is_default
                && state
                    .sheet
                    .default_style_id
                    .replace(style.id.clone())
                    .is_some()
            {
                return Err(ProjectionError::InvalidStylesXml);
            }
            if state
                .sheet
                .styles
                .insert(
                    style.id,
                    StyleDefinition {
                        based_on: (!ignored_by_word).then_some(style.based_on).flatten(),
                        properties: if ignored_by_word {
                            ParagraphProperties::default()
                        } else {
                            style.properties
                        },
                        text_properties: if ignored_by_word {
                            TextProperties::default()
                        } else {
                            style.text_properties
                        },
                        kind,
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
    frames.last().and_then(|frame| match frame {
        Frame::ParagraphProperties(owner) => Some(*owner),
        _ => None,
    })
}

fn numbering_properties_owner(frames: &[Frame]) -> Option<PropertiesOwner> {
    frames.last().and_then(|frame| match frame {
        Frame::NumberingProperties(owner) => Some(*owner),
        _ => None,
    })
}

fn run_properties_owner(frames: &[Frame]) -> Option<PropertiesOwner> {
    frames.last().and_then(|frame| match frame {
        Frame::RunProperties(owner) => Some(*owner),
        _ => None,
    })
}

fn parse_text_property(
    name: &[u8],
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    frames: &[Frame],
    sheet: &mut StyleSheet,
    current_style: &mut Option<StyleBuilder>,
) -> Result<bool, ProjectionError> {
    let Some(owner) = run_properties_owner(frames) else {
        return Ok(matches!(
            name,
            b"b" | b"bCs" | b"cs" | b"rtl" | b"highlight" | b"vertAlign"
        ));
    };
    let Some(properties) = text_properties_mut(owner, sheet, current_style) else {
        return Ok(matches!(
            name,
            b"b" | b"bCs" | b"cs" | b"rtl" | b"highlight" | b"vertAlign"
        ));
    };
    match name {
        b"b" => properties.bold = Some(on_off_element_enabled(reader, element)?),
        b"bCs" => {
            properties.complex_script_bold = Some(on_off_element_enabled(reader, element)?);
        }
        b"cs" => {
            properties.force_complex_script = Some(on_off_element_enabled(reader, element)?);
        }
        b"rtl" => properties.right_to_left = Some(on_off_element_enabled(reader, element)?),
        b"highlight" => {
            if let Some(highlighted) = attribute(reader, element, b"val")?
                .as_deref()
                .and_then(semantic_highlight_value)
            {
                properties.highlighted = Some(highlighted);
            }
        }
        b"vertAlign" => {
            if let Some(superscript) = attribute(reader, element, b"val")?
                .as_deref()
                .and_then(superscript_value)
            {
                properties.superscript = Some(superscript);
            }
        }
        _ => return Ok(false),
    }
    Ok(true)
}

fn properties_mut<'a>(
    owner: PropertiesOwner,
    sheet: &'a mut StyleSheet,
    current_style: &'a mut Option<StyleBuilder>,
) -> Option<&'a mut ParagraphProperties> {
    match owner {
        PropertiesOwner::DocumentDefaults => Some(&mut sheet.document_defaults),
        PropertiesOwner::Style => current_style.as_mut().map(|style| &mut style.properties),
        PropertiesOwner::Other => None,
    }
}

fn text_properties_mut<'a>(
    owner: PropertiesOwner,
    sheet: &'a mut StyleSheet,
    current_style: &'a mut Option<StyleBuilder>,
) -> Option<&'a mut TextProperties> {
    match owner {
        PropertiesOwner::DocumentDefaults => Some(&mut sheet.document_text_defaults),
        PropertiesOwner::Style => current_style
            .as_mut()
            .map(|style| &mut style.text_properties),
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

pub(super) fn parse_outline_level_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<u8, ProjectionError> {
    let level = parse_level_attribute(reader, element)?;
    if level > MAXIMUM_OUTLINE_LEVEL {
        return Err(ProjectionError::InvalidDocumentXml);
    }
    Ok(level)
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

fn on_off_element_enabled(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<bool, ProjectionError> {
    Ok(attribute(reader, element, b"val")?
        .is_none_or(|value| !matches!(value.to_ascii_lowercase().as_str(), "0" | "false" | "off")))
}

pub(super) const fn semantic_highlight_value(value: &str) -> Option<bool> {
    match value.as_bytes() {
        b"black" | b"blue" | b"cyan" | b"darkBlue" | b"darkCyan" | b"darkGreen"
        | b"darkMagenta" | b"darkRed" | b"darkYellow" | b"green" | b"magenta" | b"red"
        | b"yellow" => Some(true),
        b"darkGray" | b"lightGray" | b"none" | b"white" => Some(false),
        _ => None,
    }
}

#[must_use]
pub const fn is_semantic_highlight_color(value: &str) -> bool {
    matches!(semantic_highlight_value(value), Some(true))
}

const fn superscript_value(value: &str) -> Option<bool> {
    match value.as_bytes() {
        b"superscript" => Some(true),
        b"baseline" | b"subscript" => Some(false),
        _ => None,
    }
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
