use quick_xml::name::ResolveResult;

const WORDPROCESSINGML_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORDPROCESSINGML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const WORDPROCESSINGML_2010: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const MARKUP_COMPATIBILITY: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum OoxmlNamespace {
    Wordprocessing,
    Wordprocessing2010,
    MarkupCompatibility,
    Other,
}

impl OoxmlNamespace {
    pub(super) fn from_resolved(resolved: &ResolveResult<'_>) -> Self {
        let ResolveResult::Bound(namespace) = resolved else {
            return Self::Other;
        };
        match namespace.as_ref() {
            WORDPROCESSINGML_TRANSITIONAL | WORDPROCESSINGML_STRICT => Self::Wordprocessing,
            WORDPROCESSINGML_2010 => Self::Wordprocessing2010,
            MARKUP_COMPATIBILITY => Self::MarkupCompatibility,
            _ => Self::Other,
        }
    }
}

pub(super) fn is_fully_supported_namespace(namespace: &[u8]) -> bool {
    matches!(
        namespace,
        WORDPROCESSINGML_TRANSITIONAL | WORDPROCESSINGML_STRICT
    )
}
