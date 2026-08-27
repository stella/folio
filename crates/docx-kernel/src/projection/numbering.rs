use std::collections::HashMap;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use crate::ProjectionError;
use crate::projection::namespaces::OoxmlNamespace;
use crate::projection::structure::{ParagraphIndentation, StructuralFactUnknownReason};
use crate::projection::styles::parse_indentation;

const NUMBERING_LEVEL_COUNT: usize = 9;
const MAXIMUM_STYLE_LINK_DEPTH: usize = 128;

#[derive(Clone, Copy, Debug, Default)]
struct LevelDefinition {
    indentation: Option<ParagraphIndentation>,
}

#[derive(Clone, Debug)]
struct AbstractNumbering {
    levels: [Option<LevelDefinition>; NUMBERING_LEVEL_COUNT],
    num_style_link: Option<String>,
    style_link: Option<String>,
}

impl Default for AbstractNumbering {
    fn default() -> Self {
        Self {
            levels: [None; NUMBERING_LEVEL_COUNT],
            num_style_link: None,
            style_link: None,
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct LevelOverride {
    level: Option<LevelDefinition>,
}

#[derive(Clone, Debug)]
struct NumberingInstance {
    abstract_num_id: u32,
    overrides: [Option<LevelOverride>; NUMBERING_LEVEL_COUNT],
}

#[derive(Clone, Debug, Default)]
pub(super) struct NumberingCatalog {
    abstracts: HashMap<u32, AbstractNumbering>,
    instances: HashMap<u32, NumberingInstance>,
    style_links: HashMap<String, Option<u32>>,
}

impl NumberingCatalog {
    pub(super) fn indentation(
        &self,
        num_id: u32,
        level: u8,
    ) -> Result<ParagraphIndentation, StructuralFactUnknownReason> {
        let level_index = usize::from(level);
        let instance = self
            .instances
            .get(&num_id)
            .ok_or(StructuralFactUnknownReason::UnsupportedNumbering)?;
        if let Some(level_override) = instance.overrides.get(level_index).copied().flatten()
            && let Some(overridden_level) = level_override.level
        {
            return Ok(overridden_level.indentation.unwrap_or_default());
        }
        self.abstract_level(instance.abstract_num_id, level_index)
            .map(|definition| definition.indentation.unwrap_or_default())
    }

    fn abstract_level(
        &self,
        initial_abstract_num_id: u32,
        level: usize,
    ) -> Result<LevelDefinition, StructuralFactUnknownReason> {
        let initial = self
            .abstracts
            .get(&initial_abstract_num_id)
            .ok_or(StructuralFactUnknownReason::UnsupportedNumbering)?;
        if let Some(definition) = initial.levels.get(level).copied().flatten() {
            return Ok(definition);
        }
        if initial.num_style_link.is_none() {
            return Err(StructuralFactUnknownReason::UnsupportedNumbering);
        }
        self.linked_abstract_level(initial_abstract_num_id, level)
    }

    fn linked_abstract_level(
        &self,
        initial_abstract_num_id: u32,
        level: usize,
    ) -> Result<LevelDefinition, StructuralFactUnknownReason> {
        let mut current_id = initial_abstract_num_id;
        for _ in 0..MAXIMUM_STYLE_LINK_DEPTH {
            let current = self
                .abstracts
                .get(&current_id)
                .ok_or(StructuralFactUnknownReason::UnsupportedNumbering)?;
            if let Some(definition) = current.levels.get(level).copied().flatten() {
                return Ok(definition);
            }
            let Some(style_link) = current.num_style_link.as_deref() else {
                return Err(StructuralFactUnknownReason::UnsupportedNumbering);
            };
            current_id = self
                .style_links
                .get(style_link)
                .copied()
                .flatten()
                .ok_or(StructuralFactUnknownReason::UnsupportedNumbering)?;
        }
        Err(StructuralFactUnknownReason::UnsupportedNumbering)
    }
}

#[derive(Clone, Copy, Debug)]
enum Frame {
    Other,
    Numbering,
    AbstractNumbering,
    NumberingInstance,
    Level,
    LevelParagraphProperties,
    LevelOverride,
}

#[derive(Default)]
struct AbstractBuilder {
    id: u32,
    value: AbstractNumbering,
}

struct InstanceBuilder {
    id: u32,
    abstract_num_id: Option<u32>,
    overrides: [Option<LevelOverride>; NUMBERING_LEVEL_COUNT],
}

impl InstanceBuilder {
    const fn new(id: u32) -> Self {
        Self {
            id,
            abstract_num_id: None,
            overrides: [None; NUMBERING_LEVEL_COUNT],
        }
    }
}

#[derive(Clone, Copy)]
enum LevelOwner {
    Abstract,
    Override(usize),
}

#[derive(Clone, Copy)]
enum AbstractLinkKind {
    NumberingStyle,
    ParagraphStyle,
}

struct LevelBuilder {
    level: usize,
    owner: LevelOwner,
    value: LevelDefinition,
}

struct NumberingParser {
    frames: Vec<Frame>,
    catalog: NumberingCatalog,
    current_abstract: Option<AbstractBuilder>,
    current_instance: Option<InstanceBuilder>,
    current_level: Option<LevelBuilder>,
    current_override: Option<usize>,
    root_seen: bool,
    remaining_items: usize,
}

impl NumberingParser {
    fn new(maximum_items: usize) -> Self {
        Self {
            frames: Vec::new(),
            catalog: NumberingCatalog::default(),
            current_abstract: None,
            current_instance: None,
            current_level: None,
            current_override: None,
            root_seen: false,
            remaining_items: maximum_items,
        }
    }

    fn consume_item(&mut self) -> Result<(), ProjectionError> {
        self.remaining_items = self
            .remaining_items
            .checked_sub(1)
            .ok_or(ProjectionError::TooManyNumberingItems)?;
        Ok(())
    }

    fn start(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: OoxmlNamespace,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let local_name = element.local_name();
        let name = if namespace == OoxmlNamespace::Wordprocessing {
            local_name.as_ref()
        } else {
            b""
        };
        if self.frames.is_empty() && (self.root_seen || name != b"numbering") {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        let parent = self.frames.last().copied();
        let frame = match name {
            b"numbering" if self.frames.is_empty() && !self.root_seen => {
                self.root_seen = true;
                Frame::Numbering
            }
            b"abstractNum" if matches!(parent, Some(Frame::Numbering)) => {
                self.begin_abstract(reader, element)?
            }
            b"num" if matches!(parent, Some(Frame::Numbering)) => {
                self.begin_instance(reader, element)?
            }
            b"numStyleLink" if matches!(parent, Some(Frame::AbstractNumbering)) => {
                self.set_abstract_link(reader, element, AbstractLinkKind::NumberingStyle)?
            }
            b"styleLink" if matches!(parent, Some(Frame::AbstractNumbering)) => {
                self.set_abstract_link(reader, element, AbstractLinkKind::ParagraphStyle)?
            }
            b"abstractNumId" if matches!(parent, Some(Frame::NumberingInstance)) => {
                self.set_instance_abstract(reader, element)?
            }
            b"lvlOverride" if matches!(parent, Some(Frame::NumberingInstance)) => {
                self.begin_override(reader, element)?
            }
            b"lvl"
                if matches!(
                    parent,
                    Some(Frame::AbstractNumbering | Frame::LevelOverride)
                ) =>
            {
                self.begin_level(reader, element, parent)?
            }
            b"pPr" if matches!(parent, Some(Frame::Level)) => Frame::LevelParagraphProperties,
            b"ind" if matches!(parent, Some(Frame::LevelParagraphProperties)) => {
                self.set_level_indentation(reader, element)?
            }
            _ => Frame::Other,
        };
        self.frames.push(frame);
        Ok(())
    }

    fn begin_abstract(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Frame, ProjectionError> {
        if self.current_abstract.is_some() {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        self.consume_item()?;
        let id = required_u32_attribute(reader, element, b"abstractNumId")?;
        self.current_abstract = Some(AbstractBuilder {
            id,
            ..AbstractBuilder::default()
        });
        Ok(Frame::AbstractNumbering)
    }

    fn begin_instance(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Frame, ProjectionError> {
        if self.current_instance.is_some() {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        self.consume_item()?;
        self.current_instance = Some(InstanceBuilder::new(required_u32_attribute(
            reader, element, b"numId",
        )?));
        Ok(Frame::NumberingInstance)
    }

    fn set_abstract_link(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        kind: AbstractLinkKind,
    ) -> Result<Frame, ProjectionError> {
        let value = required_attribute(reader, element, b"val")?;
        let current = self
            .current_abstract
            .as_mut()
            .ok_or(ProjectionError::InvalidNumberingXml)?;
        let slot = match kind {
            AbstractLinkKind::NumberingStyle => &mut current.value.num_style_link,
            AbstractLinkKind::ParagraphStyle => &mut current.value.style_link,
        };
        if slot.replace(value).is_some() {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        Ok(Frame::Other)
    }

    fn set_instance_abstract(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Frame, ProjectionError> {
        let id = required_u32_attribute(reader, element, b"val")?;
        let current = self
            .current_instance
            .as_mut()
            .ok_or(ProjectionError::InvalidNumberingXml)?;
        if current.abstract_num_id.replace(id).is_some() {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        Ok(Frame::Other)
    }

    fn begin_override(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Frame, ProjectionError> {
        if self.current_override.is_some() {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        self.consume_item()?;
        let level = required_level_attribute(reader, element, b"ilvl")?;
        let current = self
            .current_instance
            .as_mut()
            .ok_or(ProjectionError::InvalidNumberingXml)?;
        let slot = current
            .overrides
            .get_mut(level)
            .ok_or(ProjectionError::InvalidNumberingXml)?;
        if slot.replace(LevelOverride::default()).is_some() {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        self.current_override = Some(level);
        Ok(Frame::LevelOverride)
    }

    fn begin_level(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        parent: Option<Frame>,
    ) -> Result<Frame, ProjectionError> {
        if self.current_level.is_some() {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        self.consume_item()?;
        let level = required_level_attribute(reader, element, b"ilvl")?;
        let owner = if matches!(parent, Some(Frame::AbstractNumbering)) {
            LevelOwner::Abstract
        } else {
            let override_level = self
                .current_override
                .ok_or(ProjectionError::InvalidNumberingXml)?;
            if override_level != level {
                return Err(ProjectionError::InvalidNumberingXml);
            }
            LevelOwner::Override(override_level)
        };
        self.current_level = Some(LevelBuilder {
            level,
            owner,
            value: LevelDefinition::default(),
        });
        Ok(Frame::Level)
    }

    fn set_level_indentation(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Frame, ProjectionError> {
        let indentation =
            parse_indentation(reader, element).map_err(|_| ProjectionError::InvalidNumberingXml)?;
        let current = self
            .current_level
            .as_mut()
            .ok_or(ProjectionError::InvalidNumberingXml)?;
        if current.value.indentation.replace(indentation).is_some() {
            return Err(ProjectionError::InvalidNumberingXml);
        }
        Ok(Frame::Other)
    }

    fn end(&mut self) -> Result<(), ProjectionError> {
        let frame = self
            .frames
            .pop()
            .ok_or(ProjectionError::InvalidNumberingXml)?;
        match frame {
            Frame::Level => {
                let level = self
                    .current_level
                    .take()
                    .ok_or(ProjectionError::InvalidNumberingXml)?;
                match level.owner {
                    LevelOwner::Abstract => {
                        let current = self
                            .current_abstract
                            .as_mut()
                            .ok_or(ProjectionError::InvalidNumberingXml)?;
                        let slot = current
                            .value
                            .levels
                            .get_mut(level.level)
                            .ok_or(ProjectionError::InvalidNumberingXml)?;
                        if slot.replace(level.value).is_some() {
                            return Err(ProjectionError::InvalidNumberingXml);
                        }
                    }
                    LevelOwner::Override(override_level) => {
                        let current = self
                            .current_instance
                            .as_mut()
                            .ok_or(ProjectionError::InvalidNumberingXml)?;
                        let level_override = current
                            .overrides
                            .get_mut(override_level)
                            .and_then(Option::as_mut)
                            .ok_or(ProjectionError::InvalidNumberingXml)?;
                        if level_override.level.replace(level.value).is_some() {
                            return Err(ProjectionError::InvalidNumberingXml);
                        }
                    }
                }
            }
            Frame::LevelOverride => {
                self.current_override = None;
            }
            Frame::AbstractNumbering => {
                let current = self
                    .current_abstract
                    .take()
                    .ok_or(ProjectionError::InvalidNumberingXml)?;
                if let Some(style_link) = current.value.style_link.as_ref() {
                    self.catalog
                        .style_links
                        .entry(style_link.clone())
                        .and_modify(|id| *id = None)
                        .or_insert(Some(current.id));
                }
                if self
                    .catalog
                    .abstracts
                    .insert(current.id, current.value)
                    .is_some()
                {
                    return Err(ProjectionError::InvalidNumberingXml);
                }
            }
            Frame::NumberingInstance => {
                let current = self
                    .current_instance
                    .take()
                    .ok_or(ProjectionError::InvalidNumberingXml)?;
                let abstract_num_id = current
                    .abstract_num_id
                    .ok_or(ProjectionError::InvalidNumberingXml)?;
                if self
                    .catalog
                    .instances
                    .insert(
                        current.id,
                        NumberingInstance {
                            abstract_num_id,
                            overrides: current.overrides,
                        },
                    )
                    .is_some()
                {
                    return Err(ProjectionError::InvalidNumberingXml);
                }
            }
            _ => {}
        }
        Ok(())
    }
}

pub(super) fn parse_numbering(
    xml: &[u8],
    maximum_items: usize,
) -> Result<NumberingCatalog, ProjectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut parser = NumberingParser::new(maximum_items);

    loop {
        match reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidNumberingXml)?
        {
            Event::Start(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                parser.start(&reader, OoxmlNamespace::from_resolved(&namespace), &element)?;
            }
            Event::Empty(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                parser.start(&reader, OoxmlNamespace::from_resolved(&namespace), &element)?;
                parser.end()?;
            }
            Event::End(_) => parser.end()?,
            Event::DocType(_) => return Err(ProjectionError::InvalidNumberingXml),
            Event::Eof => break,
            _ => {}
        }
    }
    if !parser.root_seen
        || !parser.frames.is_empty()
        || parser.current_abstract.is_some()
        || parser.current_instance.is_some()
        || parser.current_level.is_some()
        || parser.current_override.is_some()
    {
        return Err(ProjectionError::InvalidNumberingXml);
    }
    Ok(parser.catalog)
}

fn required_u32_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<u32, ProjectionError> {
    required_attribute(reader, element, name)?
        .parse()
        .map_err(|_| ProjectionError::InvalidNumberingXml)
}

fn required_level_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<usize, ProjectionError> {
    let level = required_u32_attribute(reader, element, name)?;
    let level = usize::try_from(level).map_err(|_| ProjectionError::InvalidNumberingXml)?;
    if level >= NUMBERING_LEVEL_COUNT {
        return Err(ProjectionError::InvalidNumberingXml);
    }
    Ok(level)
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<String, ProjectionError> {
    attribute(reader, element, name)?.ok_or(ProjectionError::InvalidNumberingXml)
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, ProjectionError> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ProjectionError::InvalidNumberingXml)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if OoxmlNamespace::from_resolved(&namespace) == OoxmlNamespace::Wordprocessing
            && local_name.as_ref() == name
        {
            if value.is_some() {
                return Err(ProjectionError::InvalidNumberingXml);
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|_| ProjectionError::InvalidNumberingXml)?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

#[cfg(test)]
#[allow(clippy::expect_used, clippy::unwrap_used)]
mod tests {
    use super::parse_numbering;
    use crate::ProjectionError;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #[test]
    fn resolves_alias_namespaces_overrides_and_style_links() {
        let xml = format!(
            r#"<n:numbering xmlns:n="{W}">
              <n:abstractNum n:abstractNumId="1"><n:styleLink n:val="Shared"/><n:lvl n:ilvl="0"><n:pPr><n:ind n:left="720"/></n:pPr></n:lvl></n:abstractNum>
              <n:abstractNum n:abstractNumId="2"><n:numStyleLink n:val="Shared"/></n:abstractNum>
              <n:num n:numId="3"><n:abstractNumId n:val="2"/></n:num>
              <n:num n:numId="4"><n:abstractNumId n:val="1"/><n:lvlOverride n:ilvl="0"><n:lvl n:ilvl="0"><n:pPr><n:ind n:left="1440" n:hanging="360"/></n:pPr></n:lvl></n:lvlOverride></n:num>
            </n:numbering>"#
        );
        let catalog = parse_numbering(xml.as_bytes(), 8).expect("numbering should parse");
        assert_eq!(catalog.indentation(3, 0).unwrap().left_twips, Some(720));
        assert_eq!(catalog.indentation(4, 0).unwrap().left_twips, Some(1440));
        assert_eq!(catalog.indentation(4, 0).unwrap().hanging_twips, Some(360));
    }

    #[test]
    fn accepts_strict_wordprocessingml_without_trusting_foreign_elements() {
        let xml = r#"<s:numbering xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:e="urn:example">
          <s:abstractNum s:abstractNumId="1"><s:lvl s:ilvl="0"><s:pPr><e:ind s:left="999"/><s:ind s:left="720"/></s:pPr></s:lvl></s:abstractNum>
          <s:num s:numId="2"><s:abstractNumId s:val="1"/></s:num>
        </s:numbering>"#;
        let catalog = parse_numbering(xml.as_bytes(), 4).expect("strict numbering should parse");
        assert_eq!(catalog.indentation(2, 0).unwrap().left_twips, Some(720));
    }

    #[test]
    fn rejects_ambiguous_invalid_and_unbounded_numbering() {
        let duplicate = format!(
            r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1"/><w:abstractNum w:abstractNumId="1"/></w:numbering>"#
        );
        assert!(matches!(
            parse_numbering(duplicate.as_bytes(), 2),
            Err(ProjectionError::InvalidNumberingXml)
        ));

        let invalid_level = format!(
            r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="9"/></w:abstractNum></w:numbering>"#
        );
        assert!(matches!(
            parse_numbering(invalid_level.as_bytes(), 2),
            Err(ProjectionError::InvalidNumberingXml)
        ));

        let multiple_roots = format!(r#"<w:numbering xmlns:w="{W}"/><w:numbering xmlns:w="{W}"/>"#);
        assert!(matches!(
            parse_numbering(multiple_roots.as_bytes(), 0),
            Err(ProjectionError::InvalidNumberingXml)
        ));

        let bounded = format!(
            r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"/></w:abstractNum></w:numbering>"#
        );
        assert!(matches!(
            parse_numbering(bounded.as_bytes(), 1),
            Err(ProjectionError::TooManyNumberingItems)
        ));

        let cyclic = format!(
            r#"<w:numbering xmlns:w="{W}">
              <w:abstractNum w:abstractNumId="1"><w:numStyleLink w:val="A"/><w:styleLink w:val="B"/></w:abstractNum>
              <w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="B"/><w:styleLink w:val="A"/></w:abstractNum>
              <w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num>
            </w:numbering>"#
        );
        let cyclic_catalog =
            parse_numbering(cyclic.as_bytes(), 4).expect("bounded cycle should parse");
        assert!(cyclic_catalog.indentation(3, 0).is_err());

        let ambiguous_link = format!(
            r#"<w:numbering xmlns:w="{W}">
              <w:abstractNum w:abstractNumId="1"><w:styleLink w:val="A"/><w:lvl w:ilvl="0"/></w:abstractNum>
              <w:abstractNum w:abstractNumId="2"><w:styleLink w:val="A"/><w:lvl w:ilvl="0"/></w:abstractNum>
              <w:abstractNum w:abstractNumId="3"><w:numStyleLink w:val="A"/></w:abstractNum>
              <w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num>
            </w:numbering>"#
        );
        let ambiguous_catalog =
            parse_numbering(ambiguous_link.as_bytes(), 6).expect("ambiguous link should parse");
        assert!(ambiguous_catalog.indentation(4, 0).is_err());
    }
}
