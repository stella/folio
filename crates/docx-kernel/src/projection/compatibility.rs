use quick_xml::XmlVersion;
use quick_xml::events::BytesStart;
use quick_xml::name::{PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;

use crate::ProjectionError;
use crate::projection::namespaces::{OoxmlNamespace, is_fully_supported_namespace};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum CompatibilityAction {
    Process,
    Skip,
}

#[derive(Default)]
struct AlternateContent {
    branch_selected: bool,
    branch_active: bool,
    fallback_seen: bool,
}

pub(super) struct MarkupCompatibility {
    alternate_content: Vec<AlternateContent>,
    suppressed_depth: usize,
    invalid_part: InvalidPart,
}

#[derive(Clone, Copy, Default)]
enum InvalidPart {
    #[default]
    Document,
    Styles,
}

impl Default for MarkupCompatibility {
    fn default() -> Self {
        Self {
            alternate_content: Vec::new(),
            suppressed_depth: 0,
            invalid_part: InvalidPart::Document,
        }
    }
}

impl MarkupCompatibility {
    pub(super) const fn for_styles() -> Self {
        Self {
            alternate_content: Vec::new(),
            suppressed_depth: 0,
            invalid_part: InvalidPart::Styles,
        }
    }

    const fn invalid_error(&self) -> ProjectionError {
        match self.invalid_part {
            InvalidPart::Document => ProjectionError::InvalidDocumentXml,
            InvalidPart::Styles => ProjectionError::InvalidStylesXml,
        }
    }

    pub(super) fn start(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: OoxmlNamespace,
        element: &BytesStart<'_>,
    ) -> Result<CompatibilityAction, ProjectionError> {
        if self.suppressed_depth > 0 {
            self.suppressed_depth = self
                .suppressed_depth
                .checked_add(1)
                .ok_or_else(|| self.invalid_error())?;
            return Ok(CompatibilityAction::Skip);
        }

        if namespace != OoxmlNamespace::MarkupCompatibility {
            if self
                .alternate_content
                .last()
                .is_some_and(|content| !content.branch_active)
            {
                self.suppressed_depth = 1;
                return Ok(CompatibilityAction::Skip);
            }
            return Ok(CompatibilityAction::Process);
        }

        match element.local_name().as_ref() {
            b"AlternateContent" => {
                if self
                    .alternate_content
                    .last()
                    .is_some_and(|content| !content.branch_active)
                {
                    return Err(self.invalid_error());
                }
                self.alternate_content.push(AlternateContent::default());
            }
            b"Choice" => {
                let supported = choice_is_supported(reader, element, self.invalid_part)?;
                self.start_branch(supported, false)?;
            }
            b"Fallback" => self.start_branch(true, true)?,
            _ => self.suppressed_depth = 1,
        }
        Ok(CompatibilityAction::Skip)
    }

    pub(super) fn empty(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: OoxmlNamespace,
        element: &BytesStart<'_>,
    ) -> Result<CompatibilityAction, ProjectionError> {
        if self.suppressed_depth > 0 {
            return Ok(CompatibilityAction::Skip);
        }
        if namespace != OoxmlNamespace::MarkupCompatibility {
            return Ok(
                if self
                    .alternate_content
                    .last()
                    .is_some_and(|content| !content.branch_active)
                {
                    CompatibilityAction::Skip
                } else {
                    CompatibilityAction::Process
                },
            );
        }

        match element.local_name().as_ref() {
            b"AlternateContent" => {
                if self
                    .alternate_content
                    .last()
                    .is_some_and(|content| !content.branch_active)
                {
                    return Err(self.invalid_error());
                }
            }
            b"Choice" => {
                let supported = choice_is_supported(reader, element, self.invalid_part)?;
                self.select_empty_branch(supported, false)?;
            }
            b"Fallback" => self.select_empty_branch(true, true)?,
            _ => {}
        }
        Ok(CompatibilityAction::Skip)
    }

    pub(super) fn end(
        &mut self,
        namespace: OoxmlNamespace,
        local_name: &[u8],
    ) -> Result<CompatibilityAction, ProjectionError> {
        if self.suppressed_depth > 0 {
            self.suppressed_depth = self
                .suppressed_depth
                .checked_sub(1)
                .ok_or_else(|| self.invalid_error())?;
            return Ok(CompatibilityAction::Skip);
        }
        if namespace != OoxmlNamespace::MarkupCompatibility {
            return Ok(CompatibilityAction::Process);
        }

        match local_name {
            b"AlternateContent" => {
                let content = self
                    .alternate_content
                    .pop()
                    .ok_or_else(|| self.invalid_error())?;
                if content.branch_active {
                    return Err(self.invalid_error());
                }
            }
            b"Choice" | b"Fallback" => {
                let invalid = self.invalid_error();
                let content = self
                    .alternate_content
                    .last_mut()
                    .ok_or_else(|| invalid.clone())?;
                if !content.branch_active {
                    return Err(invalid);
                }
                content.branch_active = false;
            }
            _ => {}
        }
        Ok(CompatibilityAction::Skip)
    }

    pub(super) const fn is_suppressed(&self) -> bool {
        self.suppressed_depth > 0
    }

    pub(super) const fn is_complete(&self) -> bool {
        self.suppressed_depth == 0 && self.alternate_content.is_empty()
    }

    fn start_branch(&mut self, supported: bool, fallback: bool) -> Result<(), ProjectionError> {
        let invalid = self.invalid_error();
        let content = self
            .alternate_content
            .last_mut()
            .ok_or_else(|| invalid.clone())?;
        validate_branch_order(content, fallback, invalid)?;
        let active = !content.branch_selected && supported;
        if active {
            content.branch_selected = true;
            content.branch_active = true;
        } else {
            self.suppressed_depth = 1;
        }
        Ok(())
    }

    fn select_empty_branch(
        &mut self,
        supported: bool,
        fallback: bool,
    ) -> Result<(), ProjectionError> {
        let invalid = self.invalid_error();
        let content = self
            .alternate_content
            .last_mut()
            .ok_or_else(|| invalid.clone())?;
        validate_branch_order(content, fallback, invalid)?;
        if !content.branch_selected && supported {
            content.branch_selected = true;
        }
        Ok(())
    }
}

const fn validate_branch_order(
    content: &mut AlternateContent,
    fallback: bool,
    invalid: ProjectionError,
) -> Result<(), ProjectionError> {
    if content.branch_active || content.fallback_seen {
        return Err(invalid);
    }
    if fallback {
        content.fallback_seen = true;
    }
    Ok(())
}

fn choice_is_supported(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    invalid: InvalidPart,
) -> Result<bool, ProjectionError> {
    let invalid = match invalid {
        InvalidPart::Document => ProjectionError::InvalidDocumentXml,
        InvalidPart::Styles => ProjectionError::InvalidStylesXml,
    };
    let mut requires = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| invalid.clone())?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if namespace == ResolveResult::Unbound && local_name.as_ref() == b"Requires" {
            if requires.is_some() {
                return Err(invalid);
            }
            requires = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|_| invalid.clone())?
                    .into_owned(),
            );
        }
    }
    let requires = requires.ok_or_else(|| invalid.clone())?;
    let mut prefixes = requires.split_ascii_whitespace().peekable();
    if prefixes.peek().is_none() {
        return Err(invalid);
    }
    Ok(prefixes.all(|required| {
        reader.resolver().bindings().any(|(prefix, namespace)| {
            matches!(prefix, PrefixDeclaration::Named(name) if name == required.as_bytes())
                && is_fully_supported_namespace(namespace.as_ref())
        })
    }))
}
