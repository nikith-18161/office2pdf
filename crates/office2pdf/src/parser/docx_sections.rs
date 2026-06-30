use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::error::ConvertWarning;
use crate::ir::{
    Block, ColumnLayout, FlowPage, HFBorder, HFBorderStyle, HFInline, HeaderFooter,
    HeaderFooterParagraph, Margins, PageSize, ParagraphStyle, Run, TextStyle,
};

use super::{
    NumberingMap, TaggedElement, extract_column_layout_from_section_property,
    extract_paragraph_style, extract_run_style, extract_tab_stop_overrides, group_into_lists,
    merge_paragraph_style, read_zip_text,
};
use crate::parser::units::twips_to_pt;

/// Parsed header/footer assets addressed by relationship ID.
#[derive(Default)]
pub(super) struct HeaderFooterAssets {
    headers: HashMap<String, HeaderFooter>,
    footers: HashMap<String, HeaderFooter>,
}

fn scan_header_footer_relationships(
    rels_xml: &str,
) -> (HashMap<String, String>, HashMap<String, String>) {
    let mut headers: HashMap<String, String> = HashMap::new();
    let mut footers: HashMap<String, String> = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(rels_xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element)) => {
                if element.local_name().as_ref() != b"Relationship" {
                    continue;
                }

                let mut id: Option<String> = None;
                let mut target: Option<String> = None;
                let mut relationship_type: Option<String> = None;

                for attribute in element.attributes().flatten() {
                    match attribute.key.local_name().as_ref() {
                        b"Id" => {
                            if let Ok(value) = attribute.unescape_value() {
                                id = Some(value.to_string());
                            }
                        }
                        b"Target" => {
                            if let Ok(value) = attribute.unescape_value() {
                                target = Some(value.to_string());
                            }
                        }
                        b"Type" => {
                            if let Ok(value) = attribute.unescape_value() {
                                relationship_type = Some(value.to_string());
                            }
                        }
                        _ => {}
                    }
                }

                let Some(id) = id else { continue };
                let Some(target) = target else { continue };
                let Some(relationship_type) = relationship_type else {
                    continue;
                };

                let full_path = if let Some(stripped) = target.strip_prefix('/') {
                    stripped.to_string()
                } else {
                    format!("word/{target}")
                };

                if relationship_type.ends_with("/header") {
                    headers.insert(id, full_path);
                } else if relationship_type.ends_with("/footer") {
                    footers.insert(id, full_path);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    (headers, footers)
}

pub(super) fn build_header_footer_assets<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> HeaderFooterAssets {
    let rels_xml = match read_zip_text(archive, "word/_rels/document.xml.rels") {
        Some(xml) => xml,
        None => return HeaderFooterAssets::default(),
    };
    let (header_relationships, footer_relationships) = scan_header_footer_relationships(&rels_xml);
    let mut assets = HeaderFooterAssets::default();

    for (relationship_id, path) in header_relationships {
        let Some(xml) = read_zip_text(archive, &path) else {
            continue;
        };
        let Ok(header) = <docx_rs::Header as docx_rs::FromXML>::from_xml(xml.as_bytes()) else {
            continue;
        };
        if let Some(converted) = convert_docx_header(&header) {
            assets.headers.insert(relationship_id, converted);
        }
    }

    for (relationship_id, path) in footer_relationships {
        let Some(xml) = read_zip_text(archive, &path) else {
            continue;
        };
        let Ok(footer) = <docx_rs::Footer as docx_rs::FromXML>::from_xml(xml.as_bytes()) else {
            continue;
        };
        if let Some(converted) = convert_docx_footer(&footer) {
            assets.footers.insert(relationship_id, converted);
        }
    }

    assets
}

pub(super) fn build_flow_page_from_section(
    section_prop: &docx_rs::SectionProperty,
    elements: Vec<TaggedElement>,
    numberings: &NumberingMap,
    header_footer_assets: &HeaderFooterAssets,
    column_layout: Option<ColumnLayout>,
    warnings: &mut Vec<ConvertWarning>,
) -> FlowPage {
    let (size, margins) = extract_page_setup(section_prop);
    let content = group_into_lists(elements, numberings);

    for block in &content {
        if let Block::Chart(chart) = block {
            let title = chart.title.as_deref().unwrap_or("untitled").to_string();
            warnings.push(ConvertWarning::FallbackUsed {
                format: "DOCX".to_string(),
                from: format!("chart ({title})"),
                to: "data table".to_string(),
            });
        }
    }

    if matches!(
        section_prop.section_type,
        Some(docx_rs::SectionType::NextColumn)
    ) {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "next-column section break".to_string(),
            to: "page-level section split".to_string(),
        });
    }

    if section_prop.first_header_reference.is_some()
        || section_prop.first_footer_reference.is_some()
        || section_prop.even_header_reference.is_some()
        || section_prop.even_footer_reference.is_some()
        || section_prop.first_header.is_some()
        || section_prop.first_footer.is_some()
        || section_prop.even_header.is_some()
        || section_prop.even_footer.is_some()
    {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "header/footer variants".to_string(),
            to: "single header/footer per section".to_string(),
        });
    }

    if section_prop
        .page_num_type
        .as_ref()
        .and_then(|page_number_type| page_number_type.start)
        .is_some()
    {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "section page number restart".to_string(),
            to: "global page counter".to_string(),
        });
    }

    FlowPage {
        size,
        margins,
        content,
        header: extract_docx_header(section_prop, header_footer_assets),
        footer: extract_docx_footer(section_prop, header_footer_assets),
        columns: column_layout
            .or_else(|| extract_column_layout_from_section_property(section_prop)),
    }
}

fn convert_docx_header(header: &docx_rs::Header) -> Option<HeaderFooter> {
    let mut paragraphs = Vec::new();
    let mut top_border: Option<HFBorder> = None;
    let mut bottom_border: Option<HFBorder> = None;
    for child in &header.children {
        match child {
            docx_rs::HeaderChild::Paragraph(paragraph) => {
                paragraphs.push(convert_hf_paragraph(paragraph));
            }
            docx_rs::HeaderChild::Table(table) => {
                let (table_paragraphs, table_top_border, table_bottom_border) =
                    convert_hf_table_to_paragraphs(table);
                paragraphs.extend(table_paragraphs);
                if top_border.is_none() {
                    top_border = table_top_border;
                }
                if bottom_border.is_none() {
                    bottom_border = table_bottom_border;
                }
            }
            docx_rs::HeaderChild::StructuredDataTag(_) => {}
        }
    }
    if paragraphs.is_empty() {
        return None;
    }
    Some(HeaderFooter {
        paragraphs,
        top_border,
        bottom_border,
    })
}
fn convert_docx_footer(footer: &docx_rs::Footer) -> Option<HeaderFooter> {
    let mut paragraphs = Vec::new();
    let mut top_border: Option<HFBorder> = None;
    let mut bottom_border: Option<HFBorder> = None;
    for child in &footer.children {
        match child {
            docx_rs::FooterChild::Paragraph(paragraph) => {
                paragraphs.push(convert_hf_paragraph(paragraph));
            }
            docx_rs::FooterChild::Table(table) => {
                let (table_paragraphs, table_top_border, table_bottom_border) =
                    convert_hf_table_to_paragraphs(table);
                paragraphs.extend(table_paragraphs);
                if top_border.is_none() {
                    top_border = table_top_border;
                }
                if bottom_border.is_none() {
                    bottom_border = table_bottom_border;
                }
            }
            docx_rs::FooterChild::StructuredDataTag(_) => {}
        }
    }
    if paragraphs.is_empty() {
        return None;
    }
    Some(HeaderFooter {
        paragraphs,
        top_border,
        bottom_border,
    })
}

/// Convert a header/footer layout table into HeaderFooterParagraphs.
///
/// Word commonly uses a borderless table to lay out left/center/right
/// content (e.g. document title | blank | "Page N") within a single
/// header/footer line. Each table row becomes one HeaderFooterParagraph,
/// with a flexible Spacer inserted between cells so the cell contents are
/// pushed apart on the line, approximating the original column layout.
fn convert_hf_table_to_paragraphs(
    table: &docx_rs::Table,
) -> (
    Vec<HeaderFooterParagraph>,
    Option<HFBorder>,
    Option<HFBorder>,
) {
    let mut paragraphs = Vec::new();
    for row_child in &table.rows {
        let docx_rs::TableChild::TableRow(row) = row_child;
        // Collect each cell's paragraphs from the row. A cell may contain
        // multiple paragraphs (e.g. "Page N" on one line, "SULIT" on the
        // next within the same footer cell); each becomes its own output
        // line rather than being concatenated onto the same line.
        let cell_paragraphs: Vec<Vec<&docx_rs::Paragraph>> = row
            .cells
            .iter()
            .map(|cell_child| {
                let docx_rs::TableRowChild::TableCell(cell) = cell_child;
                cell.children
                    .iter()
                    .filter_map(|content| match content {
                        docx_rs::TableCellContent::Paragraph(p) => Some(p),
                        _ => None,
                    })
                    .collect::<Vec<_>>()
            })
            .collect();
        let max_lines = cell_paragraphs.iter().map(Vec::len).max().unwrap_or(0);
        for line_idx in 0..max_lines {
            let mut elements: Vec<HFInline> = Vec::new();
            let mut row_style: Option<ParagraphStyle> = None;
            for (i, cell_lines) in cell_paragraphs.iter().enumerate() {
                if i > 0 {
                    elements.push(HFInline::Spacer);
                }
                if let Some(paragraph) = cell_lines.get(line_idx) {
                    if row_style.is_none() {
                        row_style = Some(extract_paragraph_style(&paragraph.property));
                    }
                    let mut field_state = HfFieldState::default();
                    collect_hf_paragraph_children(
                        &paragraph.children,
                        &mut elements,
                        &mut field_state,
                    );
                }
            }
            paragraphs.push(HeaderFooterParagraph {
                style: row_style.unwrap_or_default(),
                elements,
            });
        }
    }
    let top_border = extract_hf_table_top_border(table);
    let bottom_border = extract_hf_table_bottom_border(table);
    (paragraphs, top_border, bottom_border)
}

/// Extract a bottom border from a header layout table. Word's headers
/// commonly draw a horizontal separator BELOW the header content by
/// setting `<w:tcBorders><w:bottom .../></w:tcBorders>` on cells in the
/// last row (or every cell of a single-row layout table). The header
/// then renders as: SULIT / Document title / Section info / ─── HR ───.
/// We scan every cell in the table and use the first non-nil bottom
/// border we find — most headers have a uniform line across all cells.
fn extract_hf_table_bottom_border(table: &docx_rs::Table) -> Option<HFBorder> {
    for row_child in &table.rows {
        let docx_rs::TableChild::TableRow(row) = row_child;
        for cell_child in &row.cells {
            let docx_rs::TableRowChild::TableCell(cell) = cell_child;
            let cell_json = serde_json::to_value(&cell.property).ok()?;
            let borders = match cell_json.get("borders") {
                Some(b) => b,
                None => continue,
            };
            let bottom = match borders.get("bottom") {
                Some(b) if !b.is_null() => b,
                _ => continue,
            };
            let val = match bottom.get("borderType").and_then(|v| v.as_str()) {
                Some(v) => v,
                None => continue,
            };
            if val == "nil" || val == "none" {
                continue;
            }
            let sz_eighths = bottom.get("size").and_then(|v| v.as_u64()).unwrap_or(4);
            let thickness_pt = (sz_eighths as f64) / 8.0;
            let style = match val {
                "thickThinSmallGap" | "thinThickSmallGap" | "thickThinMediumGap"
                | "thinThickMediumGap" | "thickThinLargeGap" | "thinThickLargeGap" | "double"
                | "doubleWave" | "triple" => HFBorderStyle::Double,
                _ => HFBorderStyle::Single,
            };
            return Some(HFBorder {
                thickness_pt,
                style,
            });
        }
    }
    None
}

/// Extract the top border from a header/footer layout table, if present.
/// Word commonly draws a separator line above the footer text by setting
/// `<w:tblBorders><w:top w:val="thickThinSmallGap" w:sz="24" .../></w:tblBorders>`
/// on the layout table; the other sides are left unset and don't draw.
/// We map this back to an HFBorder so the renderer can emit a horizontal
/// rule above the header/footer paragraphs.
fn extract_hf_table_top_border(table: &docx_rs::Table) -> Option<HFBorder> {
    let property_json = serde_json::to_value(&table.property).ok()?;
    let borders_json = property_json.get("borders")?;
    let top = borders_json.get("top")?;
    let val = top.get("borderType").and_then(|v| v.as_str())?;
    if val == "nil" || val == "none" {
        return None;
    }
    // `w:sz` is in eighths of a point (1 = 0.125pt).
    let sz_eighths = top.get("size").and_then(|v| v.as_u64()).unwrap_or(4);
    let thickness_pt = (sz_eighths as f64) / 8.0;
    let style = match val {
        // Word's "thickThinSmallGap" / "thinThickSmallGap" / "double" all
        // render as parallel lines; we approximate as a double rule.
        "thickThinSmallGap" | "thinThickSmallGap" | "thickThinMediumGap" | "thinThickMediumGap"
        | "thickThinLargeGap" | "thinThickLargeGap" | "double" | "doubleWave" | "triple" => {
            HFBorderStyle::Double
        }
        _ => HFBorderStyle::Single,
    };
    Some(HFBorder {
        thickness_pt,
        style,
    })
}

/// Extract the header for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_header(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
) -> Option<HeaderFooter> {
    section_prop
        .header
        .as_ref()
        .and_then(|(_relationship_id, header)| convert_docx_header(header))
        .or_else(|| {
            section_prop
                .header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .first_header
                .as_ref()
                .and_then(|(_relationship_id, header)| convert_docx_header(header))
        })
        .or_else(|| {
            section_prop
                .first_header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .even_header
                .as_ref()
                .and_then(|(_relationship_id, header)| convert_docx_header(header))
        })
        .or_else(|| {
            section_prop
                .even_header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
}

/// Extract the footer for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_footer(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
) -> Option<HeaderFooter> {
    section_prop
        .footer
        .as_ref()
        .and_then(|(_relationship_id, footer)| convert_docx_footer(footer))
        .or_else(|| {
            section_prop
                .footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .first_footer
                .as_ref()
                .and_then(|(_relationship_id, footer)| convert_docx_footer(footer))
        })
        .or_else(|| {
            section_prop
                .first_footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .even_footer
                .as_ref()
                .and_then(|(_relationship_id, footer)| convert_docx_footer(footer))
        })
        .or_else(|| {
            section_prop
                .even_footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
}

/// Convert a docx-rs Paragraph into a HeaderFooterParagraph.
/// Detects PAGE/NUMPAGES field codes within runs and emits page counter inlines.
fn convert_hf_paragraph(paragraph: &docx_rs::Paragraph) -> HeaderFooterParagraph {
    let explicit_style = extract_paragraph_style(&paragraph.property);
    let explicit_tab_overrides = extract_tab_stop_overrides(&paragraph.property.tabs);
    let style = merge_paragraph_style(&explicit_style, explicit_tab_overrides.as_deref(), None);
    let mut elements: Vec<HFInline> = Vec::new();
    let mut field_state = HfFieldState::default();
    collect_hf_paragraph_children(&paragraph.children, &mut elements, &mut field_state);
    HeaderFooterParagraph { style, elements }
}

/// Tracks Word field-code state (begin/separate/end, and which field type is
/// active) across multiple runs within the same paragraph. A single field
/// such as `{ PAGE \* MERGEFORMAT }` is split by Word across several
/// consecutive `<w:r>` runs (one run holds fldChar begin, the next holds
/// instrText, another holds fldChar separate, another holds the cached
/// display text, and a final run holds fldChar end). Field state must
/// therefore persist across run boundaries, not reset per run.
#[derive(Default)]
struct HfFieldState {
    in_field: bool,
    past_separate: bool,
    field_inline: Option<HFInline>,
}

/// Recursively collect HFInline elements from paragraph children, descending
/// into StructuredDataTag wrappers (used by Word's quick-field building blocks
/// such as the "Page Number" field, which wraps fldChar/instrText runs in an
/// `<w:sdt>` rather than placing them as direct paragraph children).
fn collect_hf_paragraph_children(
    children: &[docx_rs::ParagraphChild],
    elements: &mut Vec<HFInline>,
    field_state: &mut HfFieldState,
) {
    for child in children {
        match child {
            docx_rs::ParagraphChild::Run(run) => {
                let run_style = extract_run_style(&run.run_property);
                extract_hf_run_elements(&run.children, &run_style, elements, field_state);
            }
            docx_rs::ParagraphChild::StructuredDataTag(sdt) => {
                collect_hf_sdt_children(&sdt.children, elements, field_state);
            }
            _ => {}
        }
    }
}

/// Recursively collect HFInline elements from StructuredDataTag children.
fn collect_hf_sdt_children(
    children: &[docx_rs::StructuredDataTagChild],
    elements: &mut Vec<HFInline>,
    field_state: &mut HfFieldState,
) {
    for child in children {
        match child {
            docx_rs::StructuredDataTagChild::Run(run) => {
                let run_style = extract_run_style(&run.run_property);
                extract_hf_run_elements(&run.children, &run_style, elements, field_state);
            }
            docx_rs::StructuredDataTagChild::Paragraph(para) => {
                collect_hf_paragraph_children(&para.children, elements, field_state);
            }
            docx_rs::StructuredDataTagChild::StructuredDataTag(nested) => {
                collect_hf_sdt_children(&nested.children, elements, field_state);
            }
            _ => {}
        }
    }
}

/// Extract inline elements from a run's children for header/footer use.
/// Recognizes text, tabs, and PAGE/NUMPAGES field codes. Field state is
/// threaded in via `field_state` so that begin/separate/end markers split
/// across multiple runs (the common case for Word's quick-field building
/// blocks) are tracked correctly instead of resetting on every run.
fn extract_hf_run_elements(
    children: &[docx_rs::RunChild],
    style: &TextStyle,
    elements: &mut Vec<HFInline>,
    field_state: &mut HfFieldState,
) {
    for child in children {
        match child {
            docx_rs::RunChild::FieldChar(field_char) => match field_char.field_char_type {
                docx_rs::FieldCharType::Begin => {
                    field_state.in_field = true;
                    field_state.field_inline = None;
                    field_state.past_separate = false;
                }
                docx_rs::FieldCharType::Separate => {
                    field_state.past_separate = true;
                }
                docx_rs::FieldCharType::End => {
                    if let Some(inline) = field_state.field_inline.take() {
                        elements.push(inline);
                    }
                    field_state.in_field = false;
                    field_state.past_separate = false;
                }
                _ => {}
            },
            docx_rs::RunChild::InstrText(instruction) => {
                if !field_state.in_field {
                    continue;
                }
                field_state.field_inline = match instruction.as_ref() {
                    docx_rs::InstrText::PAGE(_) => Some(HFInline::PageNumber),
                    docx_rs::InstrText::NUMPAGES(_) => Some(HFInline::TotalPages),
                    _ => field_state.field_inline.take(),
                };
            }
            docx_rs::RunChild::InstrTextString(value) => {
                if !field_state.in_field {
                    continue;
                }
                // Field instructions look like " PAGE   \* MERGEFORMAT " or
                // " NUMPAGES \* MERGEFORMAT " — the field keyword is the
                // first whitespace-separated token, followed by switches
                // such as `\* MERGEFORMAT`. Match on that leading token
                // rather than the whole trimmed string.
                let keyword = value.trim().split_whitespace().next().unwrap_or("");
                if keyword.eq_ignore_ascii_case("page") {
                    field_state.field_inline = Some(HFInline::PageNumber);
                } else if keyword.eq_ignore_ascii_case("numpages") {
                    field_state.field_inline = Some(HFInline::TotalPages);
                }
            }
            docx_rs::RunChild::Text(text) => {
                if field_state.in_field && field_state.past_separate {
                    continue;
                }
                if !field_state.in_field && !text.text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: text.text.clone(),
                        style: style.clone(),
                        href: None,
                        footnote: None,
                    }));
                }
            }
            docx_rs::RunChild::Tab(_) if !field_state.in_field => {
                elements.push(HFInline::Run(Run {
                    text: "\t".to_string(),
                    style: style.clone(),
                    href: None,
                    footnote: None,
                }));
            }
            _ => {}
        }
    }
}
/// Extract page size and margins from DOCX section properties.
fn extract_page_setup(section_prop: &docx_rs::SectionProperty) -> (PageSize, Margins) {
    let size = extract_page_size(&section_prop.page_size);
    let margins = extract_margins(&section_prop.page_margin);
    (size, margins)
}

fn round_page_dimension_pt(value: f64) -> f64 {
    (value * 2.0).round() / 2.0
}

/// Extract page size from docx-rs PageSize (which has private fields).
/// Uses serde serialization to access the private `w`, `h`, and `orient` fields.
/// Values in DOCX are in twips (1/20 of a point).
/// When orient is "landscape" and width < height, dimensions are swapped to ensure
/// landscape pages have width > height.
pub(super) fn extract_page_size(page_size: &docx_rs::PageSize) -> PageSize {
    if let Ok(json) = serde_json::to_value(page_size) {
        let width_twips = json
            .get("w")
            .and_then(|value| value.as_f64())
            .unwrap_or(0.0);
        let height_twips = json
            .get("h")
            .and_then(|value| value.as_f64())
            .unwrap_or(0.0);
        let orientation = json.get("orient").and_then(|value| value.as_str());
        if width_twips > 0.0 && height_twips > 0.0 {
            let mut width = round_page_dimension_pt(twips_to_pt(width_twips));
            let mut height = round_page_dimension_pt(twips_to_pt(height_twips));
            if orientation == Some("landscape") && width < height {
                std::mem::swap(&mut width, &mut height);
            }
            return PageSize { width, height };
        }
    }
    PageSize::default()
}

/// Extract margins from docx-rs PageMargin.
/// PageMargin fields are public i32 values in twips.
fn extract_margins(page_margin: &docx_rs::PageMargin) -> Margins {
    Margins {
        top: twips_to_pt(page_margin.top),
        bottom: twips_to_pt(page_margin.bottom),
        left: twips_to_pt(page_margin.left),
        right: twips_to_pt(page_margin.right),
    }
}
