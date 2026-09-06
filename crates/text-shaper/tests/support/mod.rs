//! WOFF 1.0 to plain sfnt, for tests only.
//!
//! The bundled faces ship as WOFF and a shaper reads sfnt, so a test that
//! skipped on the container would assert nothing at all. Production callers
//! hand this crate sfnt bytes; the TypeScript side already owns that decode.

#![allow(clippy::redundant_pub_crate, clippy::missing_const_for_fn)]

use miniz_oxide::inflate::decompress_to_vec_zlib;

const HEADER_LENGTH: usize = 44;
const DIRECTORY_ENTRY_LENGTH: usize = 20;
const SFNT_HEADER_LENGTH: usize = 12;
const SFNT_ENTRY_LENGTH: usize = 16;

fn u32_at(bytes: &[u8], offset: usize) -> Option<u32> {
    let end = offset.checked_add(4)?;
    let slice = bytes.get(offset..end)?;
    Some(u32::from_be_bytes(slice.try_into().ok()?))
}

fn u16_at(bytes: &[u8], offset: usize) -> Option<u16> {
    let end = offset.checked_add(2)?;
    let slice = bytes.get(offset..end)?;
    Some(u16::from_be_bytes(slice.try_into().ok()?))
}

/// Returns plain sfnt bytes, or the input unchanged when it is already sfnt.
pub(crate) fn to_sfnt(bytes: &[u8]) -> Option<Vec<u8>> {
    if bytes.get(0..4) != Some(b"wOFF") {
        return Some(bytes.to_vec());
    }
    let flavor = u32_at(bytes, 4)?;
    let table_count = usize::from(u16_at(bytes, 12)?);

    let mut tables = Vec::with_capacity(table_count);
    for index in 0..table_count {
        let entry = HEADER_LENGTH.checked_add(index.checked_mul(DIRECTORY_ENTRY_LENGTH)?)?;
        let tag = bytes.get(entry..entry.checked_add(4)?)?.to_vec();
        let offset = usize::try_from(u32_at(bytes, entry.checked_add(4)?)?).ok()?;
        let compressed = usize::try_from(u32_at(bytes, entry.checked_add(8)?)?).ok()?;
        let original = usize::try_from(u32_at(bytes, entry.checked_add(12)?)?).ok()?;
        let span = bytes.get(offset..offset.checked_add(compressed)?)?;
        let data = if compressed == original {
            span.to_vec()
        } else {
            decompress_to_vec_zlib(span).ok()?
        };
        tables.push((tag, data));
    }
    tables.sort_by(|left, right| left.0.cmp(&right.0));

    let mut out = Vec::new();
    out.extend_from_slice(&flavor.to_be_bytes());
    let count = u16::try_from(tables.len()).ok()?;
    out.extend_from_slice(&count.to_be_bytes());
    // searchRange, entrySelector and rangeShift: derived, and no reader this
    // is handed to depends on them being anything but self-consistent.
    let highest_bit = usize::BITS
        .saturating_sub(1)
        .saturating_sub(tables.len().leading_zeros());
    let entry_selector = u16::try_from(highest_bit).unwrap_or(0);
    let search_range = 1_u16
        .checked_shl(u32::from(entry_selector))
        .unwrap_or(1)
        .saturating_mul(16);
    out.extend_from_slice(&search_range.to_be_bytes());
    out.extend_from_slice(&entry_selector.to_be_bytes());
    out.extend_from_slice(
        &count
            .saturating_mul(16)
            .saturating_sub(search_range)
            .to_be_bytes(),
    );

    let mut offset =
        SFNT_HEADER_LENGTH.checked_add(tables.len().checked_mul(SFNT_ENTRY_LENGTH)?)?;
    let mut body = Vec::new();
    for (tag, data) in &tables {
        out.extend_from_slice(tag);
        out.extend_from_slice(&0_u32.to_be_bytes());
        out.extend_from_slice(&u32::try_from(offset).ok()?.to_be_bytes());
        out.extend_from_slice(&u32::try_from(data.len()).ok()?.to_be_bytes());
        body.extend_from_slice(data);
        let padding = data.len().next_multiple_of(4).saturating_sub(data.len());
        body.extend(std::iter::repeat_n(0_u8, padding));
        offset = offset.checked_add(data.len())?.checked_add(padding)?;
    }
    out.extend_from_slice(&body);
    Some(out)
}
