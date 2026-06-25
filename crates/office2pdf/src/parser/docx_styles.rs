use std::collections::HashMap;

use crate::ir::{ParagraphStyle, TabStop, TextStyle};

use super::text::{
    extract_doc_default_paragraph_style, extract_doc_default_text_style, extract_paragraph_style,
    extract_run_style, extract_tab_stop_overrides,
};

/// Resolved style formatting extracted from a document style definition.
/// Contains text and paragraph formatting along with an optional heading level.
pub(super) struct ResolvedStyle {
    pub(super) text: TextStyle,
    pub(super) paragraph: ParagraphStyle,
    pub(super) paragraph_tab_overrides: Option<Vec<TabStopOverride>>,
    /// Heading level from outline_lvl (0 = Heading 1, 1 = Heading 2, ..., 5 = Heading 6).
    pub(super) heading_level: Option<usize>,
    /// Numbering info read from the style's own <w:numPr> (i.e. inherited by
    /// any paragraph that uses this style). This is how Word's built-in
    /// Heading1-9 styles point at a multilevel numbering scheme (numId=1,
    /// ilvl=0..3) — heading paragraphs in the body don't carry an inline
    /// <w:numPr>, they inherit it from the style. None means the style has
    /// no numbering attached, so paragraphs using it render unnumbered.
    pub(super) style_num_info: Option<super::lists::NumInfo>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) enum TabStopOverride {
    Set(TabStop),
    Clear(f64),
}

/// Map from style_id → resolved formatting.
pub(super) type StyleMap = HashMap<String, ResolvedStyle>;

/// Synthetic style ID used for document-level default text properties.
pub(super) const DOC_DEFAULT_STYLE_ID: &str = "__office2pdf_doc_defaults";

use crate::defaults::HEADING_FONT_SIZES;

/// Build a map from style ID → resolved formatting by extracting formatting
/// from each style's run_property and paragraph_property.
/// Extract <w:numPr> from a docx-rs ParagraphProperty if present.
/// Returns None when the style has no numbering reference or numId == 0.
fn extract_style_num_info(prop: &docx_rs::ParagraphProperty) -> Option<super::lists::NumInfo> {
    let np = prop.numbering_property.as_ref()?;
    let num_id = np.id.as_ref()?.id;
    if num_id == 0 {
        return None;
    }
    let level = np.level.as_ref().map_or(0, |l| l.val as u32);
    Some(super::lists::NumInfo { num_id, level })
}

pub(super) fn build_style_map(styles: &docx_rs::Styles) -> StyleMap {
    let mut map = StyleMap::new();
    let default_text: TextStyle = extract_doc_default_text_style(styles);
    let default_paragraph: ParagraphStyle = extract_doc_default_paragraph_style(styles);

    map.insert(
        DOC_DEFAULT_STYLE_ID.to_string(),
        ResolvedStyle {
            text: default_text,
            paragraph: default_paragraph,
            paragraph_tab_overrides: None,
            heading_level: None,
            style_num_info: None,
        },
    );

    for style in &styles.styles {
        match style.style_type {
            docx_rs::StyleType::Paragraph => {
                let text = merge_text_style(
                    &extract_run_style(&style.run_property),
                    map.get(DOC_DEFAULT_STYLE_ID),
                );
                let paragraph_tab_overrides =
                    extract_tab_stop_overrides(&style.paragraph_property.tabs);
                let paragraph = merge_paragraph_style(
                    &extract_paragraph_style(&style.paragraph_property),
                    paragraph_tab_overrides.as_deref(),
                    map.get(DOC_DEFAULT_STYLE_ID),
                );
                let heading_level = style
                    .paragraph_property
                    .outline_lvl
                    .as_ref()
                    .map(|outline_level| outline_level.v)
                    .filter(|&value| value < 6);

                let style_num_info = extract_style_num_info(&style.paragraph_property);
                map.insert(
                    style.style_id.clone(),
                    ResolvedStyle {
                        text,
                        paragraph,
                        paragraph_tab_overrides,
                        heading_level,
                        style_num_info,
                    },
                );
            }
            // Character styles (e.g. pandoc's `BuiltInTok`/`StringTok` syntax
            // highlighting tokens) contribute only run-level text properties.
            // They deliberately do NOT inherit document defaults, so that
            // overlaying a run's `rStyle` onto its paragraph style changes only
            // the properties the character style actually sets (issue #176).
            docx_rs::StyleType::Character => {
                let mut text = extract_run_style(&style.run_property);
                // Word's built-in link character styles ("Hyperlink",
                // "FollowedHyperlink", "InternetLink") set blue text and a
                // single underline. Those properties are an *on-screen*
                // affordance in Word: when the same document is exported
                // to PDF, Word drops them and renders hyperlinks (including
                // TOC PAGEREF anchors) using the surrounding paragraph's
                // color. Our renderer is a PDF converter, not an editor
                // surface, so we mirror Word's PDF-export behaviour and
                // strip color + underline from these specific styleIds.
                // Other properties the styles set (rare in practice) pass
                // through. This is safer than special-casing TOC
                // paragraphs because Word's suppression applies to every
                // run that references one of these built-in link styles,
                // not just TOC entries — and the alternative (threading
                // hyperlink-kind context all the way down to run
                // extraction) would change several signatures for the
                // sake of a single cosmetic rule.
                if matches!(
                    style.style_id.as_str(),
                    "Hyperlink" | "FollowedHyperlink" | "InternetLink"
                ) {
                    text.color = None;
                    text.underline = None;
                }
                map.insert(
                    style.style_id.clone(),
                    ResolvedStyle {
                        text,
                        paragraph: ParagraphStyle::default(),
                        paragraph_tab_overrides: None,
                        heading_level: None,
                        style_num_info: None,
                    },
                );
            }
            _ => {}
        }
    }

    map
}

/// Merge style text formatting with explicit run formatting.
/// Explicit formatting (from the run itself) takes priority over style formatting.
/// For heading styles, default sizes and bold are applied when neither the style
/// nor the run specifies them.
pub(super) fn merge_text_style(explicit: &TextStyle, style: Option<&ResolvedStyle>) -> TextStyle {
    let (style_text, heading_level) = match style {
        Some(style) => (&style.text, style.heading_level),
        None => return explicit.clone(),
    };

    let mut merged: TextStyle = style_text.clone();

    // Heading defaults: apply fallback size/bold when the style itself
    // doesn't specify them. This must happen before the explicit overwrite
    // so that explicit values still win.
    if let Some(level) = heading_level {
        if merged.font_size.is_none() {
            merged.font_size = Some(HEADING_FONT_SIZES[level]);
        }
        if merged.bold.is_none() {
            merged.bold = Some(true);
        }
    }

    merged.merge_from(explicit);

    merged
}

/// Merge style paragraph formatting with explicit paragraph formatting.
/// Explicit formatting takes priority.
pub(super) fn merge_paragraph_style(
    explicit: &ParagraphStyle,
    explicit_tab_overrides: Option<&[TabStopOverride]>,
    style: Option<&ResolvedStyle>,
) -> ParagraphStyle {
    let style_paragraph = style.map(|resolved_style| &resolved_style.paragraph);
    let inherited_tab_stops = style.and_then(resolve_style_tab_stops);

    ParagraphStyle {
        alignment: explicit
            .alignment
            .or(style_paragraph.and_then(|style| style.alignment)),
        indent_left: explicit
            .indent_left
            .or(style_paragraph.and_then(|style| style.indent_left)),
        indent_right: explicit
            .indent_right
            .or(style_paragraph.and_then(|style| style.indent_right)),
        indent_first_line: explicit
            .indent_first_line
            .or(style_paragraph.and_then(|style| style.indent_first_line)),
        font_size: explicit
            .font_size
            .or(style_paragraph.and_then(|style| style.font_size)),
        line_spacing: explicit
            .line_spacing
            .or(style_paragraph.and_then(|style| style.line_spacing)),
        space_before: explicit
            .space_before
            .or(style_paragraph.and_then(|style| style.space_before)),
        space_after: explicit
            .space_after
            .or(style_paragraph.and_then(|style| style.space_after)),
        heading_level: style
            .and_then(|resolved_style| resolved_style.heading_level)
            .map(|level| (level + 1) as u8),
        heading_number: None,
        direction: explicit.direction,
        tab_stops: merge_tab_stops(
            explicit.tab_stops.as_deref(),
            explicit_tab_overrides,
            inherited_tab_stops.as_deref(),
        ),
    }
}

fn resolve_style_tab_stops(style: &ResolvedStyle) -> Option<Vec<TabStop>> {
    resolve_tab_stop_source(
        style.paragraph.tab_stops.as_deref(),
        style.paragraph_tab_overrides.as_deref(),
    )
}

fn resolve_tab_stop_source(
    tab_stops: Option<&[TabStop]>,
    tab_overrides: Option<&[TabStopOverride]>,
) -> Option<Vec<TabStop>> {
    if let Some(tab_overrides) = tab_overrides {
        let mut resolved: Vec<TabStop> = Vec::new();
        apply_tab_stop_overrides(&mut resolved, tab_overrides);
        return Some(resolved);
    }

    tab_stops.map(|tab_stops| tab_stops.to_vec())
}

fn merge_tab_stops(
    explicit_tab_stops: Option<&[TabStop]>,
    explicit_tab_overrides: Option<&[TabStopOverride]>,
    inherited_tab_stops: Option<&[TabStop]>,
) -> Option<Vec<TabStop>> {
    if let Some(explicit_tab_overrides) = explicit_tab_overrides {
        let mut resolved: Vec<TabStop> = inherited_tab_stops.unwrap_or(&[]).to_vec();
        apply_tab_stop_overrides(&mut resolved, explicit_tab_overrides);
        return Some(resolved);
    }

    explicit_tab_stops
        .map(|tab_stops| tab_stops.to_vec())
        .or_else(|| inherited_tab_stops.map(|tab_stops| tab_stops.to_vec()))
}

pub(super) fn apply_tab_stop_overrides(
    tab_stops: &mut Vec<TabStop>,
    tab_overrides: &[TabStopOverride],
) {
    for tab_override in tab_overrides {
        match tab_override {
            TabStopOverride::Set(tab_stop) => {
                tab_stops.retain(|existing| {
                    !tab_stop_positions_match(existing.position, tab_stop.position)
                });
                tab_stops.push(*tab_stop);
            }
            TabStopOverride::Clear(position) => {
                tab_stops
                    .retain(|existing| !tab_stop_positions_match(existing.position, *position));
            }
        }
    }

    tab_stops.sort_by(|left, right| {
        left.position
            .partial_cmp(&right.position)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
}

fn tab_stop_positions_match(left: f64, right: f64) -> bool {
    (left - right).abs() < 0.01
}

/// Look up the pStyle reference from a paragraph's property.
pub(super) fn get_paragraph_style_id(prop: &docx_rs::ParagraphProperty) -> Option<&str> {
    prop.style.as_ref().map(|style| style.val.as_str())
}
