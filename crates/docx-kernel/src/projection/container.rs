//! Paragraph containers emitted by the current projection producer.

macro_rules! paragraph_containers {
    ($first:ident => $first_wire:literal $(, $variant:ident => $wire:literal)* $(,)?) => {
        #[derive(Clone, Copy, Debug, Eq, PartialEq)]
        pub enum ParagraphContainer {
            $first,
            $($variant,)*
        }

        #[cfg(all(target_arch = "wasm32", feature = "wasm"))]
        impl ParagraphContainer {
            pub(crate) const fn wire_name(self) -> &'static str {
                match self {
                    Self::$first => $first_wire,
                    $(Self::$variant => $wire,)*
                }
            }
        }

        #[cfg(all(target_arch = "wasm32", feature = "wasm"))]
        #[wasm_bindgen::prelude::wasm_bindgen(typescript_custom_section)]
        const TYPESCRIPT_CONTAINER: &str = concat!(
            "export type DocxProjectionContainer = ",
            stringify!($first_wire),
            $(" | ", stringify!($wire),)*
            ";",
        );
    };
}

// Textboxes and non-body stories remain outside this projection producer.
paragraph_containers!(Body => "body", TableCell => "tableCell");
