//! Raw-XML side-channel for floating drawing *shapes* (`wps:wsp`).
//!
//! docx-rs (the upstream DOCX parser) only models `<w:drawing>` data as either a
//! picture (`Pic`) or a text box (`TextBox`). A DrawingML word-processing shape
//! that carries geometry but no text box — e.g. a `<a:prstGeom prst="rect">`
//! rectangle or a `prst="line"` connector/arrow authored by LibreOffice — parses
//! into a `Drawing` with `data == None`, so its geometry, fill and stroke are
//! lost entirely (issue #176).
//!
//! This module scans the raw `word/document.xml` for such shapes in document
//! order and exposes them through a cursor that the main walk advances once per
//! geometry-only drawing it encounters, mirroring [`DrawingTextBoxContext`] and
//! [`VmlTextBoxContext`].
//!
//! [`DrawingTextBoxContext`]: super::docx_context_drawing::DrawingTextBoxContext
//! [`VmlTextBoxContext`]: super::docx_context_vml::VmlTextBoxContext

use std::cell::Cell;

use quick_xml::events::{BytesStart, Event};

use crate::ir::{
    ArrowHead, BorderLineStyle, BorderSide, Color, FloatingShape, Shape, ShapeKind, WrapMode,
};
use crate::parser::units::emu_to_pt;
use crate::parser::xml_util::parse_hex_color;

/// Default stroke width (pt) when a shape's `<a:ln w="0">` requests the
/// renderer's hairline default. Word/LibreOffice treat `w="0"` as "thin but
/// visible"; 0 pt would make the outline disappear.
const DEFAULT_STROKE_WIDTH_PT: f64 = 0.75;

/// EMU per point (914400 EMU/inch ÷ 72 pt/inch).
const EMU_PER_POINT: f64 = 12700.0;

/// Floating geometry-only shapes scanned from `word/document.xml`, consumed in
/// document order alongside the docx-rs element walk.
#[derive(Debug, Clone)]
pub(in super::super) struct DrawingShapeContext {
    shapes: Vec<FloatingShape>,
    cursor: Cell<usize>,
}

impl DrawingShapeContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            shapes: xml.map(scan_drawing_shapes).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    /// Return the next scanned shape, advancing the cursor. Returns `None` once
    /// the scanned shapes are exhausted so a mismatched walk degrades to "no
    /// shape" rather than panicking.
    pub(in super::super) fn consume_next(&self) -> Option<FloatingShape> {
        let index: usize = self.cursor.get();
        self.cursor.set(index + 1);
        self.shapes.get(index).cloned()
    }
}

/// Which `<wp:positionH>` / `<wp:positionV>` axis the current `<wp:posOffset>`
/// text belongs to.
#[derive(Clone, Copy, PartialEq, Eq)]
enum PositionAxis {
    None,
    Horizontal,
    Vertical,
}

/// Mutable accumulator for a single `<w:drawing>` while scanning.
#[derive(Default)]
struct ShapeBuilder {
    has_wsp: bool,
    has_text_box: bool,
    preset: Option<String>,
    box_width_pt: Option<f64>,
    box_height_pt: Option<f64>,
    offset_x_pt: f64,
    offset_y_pt: f64,
    flip_h: bool,
    flip_v: bool,
    fill_color: Option<Color>,
    fill_none: bool,
    line_color: Option<Color>,
    line_width_pt: Option<f64>,
    line_none: bool,
    has_line: bool,
    head_arrow: bool,
    tail_arrow: bool,
}

impl ShapeBuilder {
    /// Build a [`FloatingShape`] from the accumulated geometry, or `None` when
    /// this drawing is not a geometry-only shape (it is a picture or a text box,
    /// both handled by docx-rs).
    fn finish(self) -> Option<FloatingShape> {
        if !self.has_wsp || self.has_text_box {
            return None;
        }

        let width: f64 = self.box_width_pt.unwrap_or(0.0);
        let height: f64 = self.box_height_pt.unwrap_or(0.0);
        let kind: ShapeKind = self.resolve_kind(width, height);

        let fill: Option<Color> = if self.fill_none {
            None
        } else {
            self.fill_color
        };
        let stroke: Option<BorderSide> = self.resolve_stroke();

        // A shape with neither fill, stroke nor a line geometry would render as
        // nothing — skip it so we stay in sync with the renderer.
        let is_line: bool = matches!(kind, ShapeKind::Line { .. });
        if fill.is_none() && stroke.is_none() && !is_line {
            return None;
        }

        Some(FloatingShape {
            shape: Shape {
                kind,
                fill,
                gradient_fill: None,
                stroke,
                rotation_deg: None,
                opacity: None,
                shadow: None,
            },
            width,
            height,
            offset_x: self.offset_x_pt,
            offset_y: self.offset_y_pt,
            wrap_mode: WrapMode::None,
        })
    }

    fn resolve_kind(&self, width: f64, height: f64) -> ShapeKind {
        match self.preset.as_deref() {
            Some("line") | Some("straightConnector1") => {
                // Endpoints run corner-to-corner of the bounding box; flips swap
                // the diagonal direction (no-op for axis-aligned lines).
                let (x1, x2) = if self.flip_h {
                    (width, 0.0)
                } else {
                    (0.0, width)
                };
                let (y1, y2) = if self.flip_v {
                    (height, 0.0)
                } else {
                    (0.0, height)
                };
                ShapeKind::Line {
                    x1,
                    y1,
                    x2,
                    y2,
                    head_end: arrow(self.head_arrow),
                    tail_end: arrow(self.tail_arrow),
                }
            }
            Some("ellipse") | Some("oval") => ShapeKind::Ellipse,
            Some("roundRect") => ShapeKind::RoundedRectangle {
                radius_fraction: 0.1,
            },
            // "rect" and any unsupported preset fall back to a rectangle so the
            // shape's area, fill and outline are still conveyed.
            _ => ShapeKind::Rectangle,
        }
    }

    fn resolve_stroke(&self) -> Option<BorderSide> {
        if self.line_none || !self.has_line {
            return None;
        }
        let width: f64 = match self.line_width_pt {
            Some(width) if width > 0.0 => width,
            _ => DEFAULT_STROKE_WIDTH_PT,
        };
        Some(BorderSide {
            width,
            color: self.line_color.unwrap_or(Color { r: 0, g: 0, b: 0 }),
            style: BorderLineStyle::Solid,
        })
    }
}

fn arrow(present: bool) -> ArrowHead {
    if present {
        ArrowHead::Triangle
    } else {
        ArrowHead::None
    }
}

fn attribute_value(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        (attribute.key.local_name().as_ref() == name)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

fn emu_attr_to_pt(element: &BytesStart<'_>, name: &[u8]) -> Option<f64> {
    attribute_value(element, name)
        .and_then(|value| value.parse::<i64>().ok())
        .map(emu_to_pt)
}

fn bool_attr(element: &BytesStart<'_>, name: &[u8]) -> bool {
    matches!(
        attribute_value(element, name).as_deref(),
        Some("1") | Some("true")
    )
}

/// Scan `word/document.xml`, returning one [`FloatingShape`] per geometry-only
/// `wps:wsp` drawing, in document order.
fn scan_drawing_shapes(xml: &str) -> Vec<FloatingShape> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut result: Vec<FloatingShape> = Vec::new();

    let mut drawing_depth: usize = 0;
    let mut sppr_depth: usize = 0;
    let mut line_depth: usize = 0;
    let mut axis: PositionAxis = PositionAxis::None;
    let mut in_position_offset: bool = false;
    let mut builder: Option<ShapeBuilder> = None;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref element)) => match element.local_name().as_ref() {
                b"drawing" => {
                    drawing_depth += 1;
                    if drawing_depth == 1 {
                        builder = Some(ShapeBuilder::default());
                        sppr_depth = 0;
                        line_depth = 0;
                        axis = PositionAxis::None;
                        in_position_offset = false;
                    }
                }
                b"positionH" => axis = PositionAxis::Horizontal,
                b"positionV" => axis = PositionAxis::Vertical,
                b"posOffset" => in_position_offset = true,
                b"spPr" if builder.is_some() => sppr_depth += 1,
                b"ln" if builder.is_some() => {
                    line_depth += 1;
                    if let Some(builder) = builder.as_mut() {
                        builder.has_line = true;
                        builder.line_width_pt =
                            emu_attr_to_pt(element, b"w").or(builder.line_width_pt);
                    }
                }
                other => handle_geometry_element(
                    builder.as_mut(),
                    other,
                    element,
                    sppr_depth,
                    line_depth,
                ),
            },
            Ok(Event::Empty(ref element)) => {
                handle_geometry_element(
                    builder.as_mut(),
                    element.local_name().as_ref(),
                    element,
                    sppr_depth,
                    line_depth,
                );
            }
            Ok(Event::Text(ref text)) => {
                if in_position_offset
                    && let Some(builder) = builder.as_mut()
                    && let Ok(raw) = text.xml_content()
                    && let Ok(emu) = raw.trim().parse::<i64>()
                {
                    let pt: f64 = (emu as f64) / EMU_PER_POINT;
                    match axis {
                        PositionAxis::Horizontal => builder.offset_x_pt = pt,
                        PositionAxis::Vertical => builder.offset_y_pt = pt,
                        PositionAxis::None => {}
                    }
                }
            }
            Ok(Event::End(ref element)) => match element.local_name().as_ref() {
                b"posOffset" => in_position_offset = false,
                b"positionH" | b"positionV" => axis = PositionAxis::None,
                b"spPr" if sppr_depth > 0 => sppr_depth -= 1,
                b"ln" if line_depth > 0 => line_depth -= 1,
                b"drawing" if drawing_depth > 0 => {
                    drawing_depth -= 1;
                    if drawing_depth == 0
                        && let Some(shape) = builder.take().and_then(ShapeBuilder::finish)
                    {
                        result.push(shape);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    result
}

/// Apply a geometry/fill/stroke element (`wsp`, `txbx`, `extent`, `prstGeom`,
/// `xfrm`, `srgbClr`, `noFill`, `tailEnd`, `headEnd`) to the current builder.
fn handle_geometry_element(
    builder: Option<&mut ShapeBuilder>,
    local_name: &[u8],
    element: &BytesStart<'_>,
    sppr_depth: usize,
    line_depth: usize,
) {
    let Some(builder) = builder else {
        return;
    };

    match local_name {
        b"wsp" => builder.has_wsp = true,
        b"txbx" => builder.has_text_box = true,
        // The anchor extent gives the on-page bounding box.
        b"extent" => {
            if let Some(width) = emu_attr_to_pt(element, b"cx") {
                builder.box_width_pt = Some(width);
            }
            if let Some(height) = emu_attr_to_pt(element, b"cy") {
                builder.box_height_pt = Some(height);
            }
        }
        b"prstGeom" => {
            if let Some(preset) = attribute_value(element, b"prst") {
                builder.preset = Some(preset);
            }
        }
        b"xfrm" => {
            builder.flip_h = bool_attr(element, b"flipH");
            builder.flip_v = bool_attr(element, b"flipV");
        }
        b"srgbClr" if sppr_depth > 0 => {
            if let Some(color) = attribute_value(element, b"val").and_then(|v| parse_hex_color(&v))
            {
                if line_depth > 0 {
                    builder.line_color = builder.line_color.or(Some(color));
                } else {
                    builder.fill_color = builder.fill_color.or(Some(color));
                }
            }
        }
        b"noFill" if sppr_depth > 0 => {
            if line_depth > 0 {
                builder.line_none = true;
            } else {
                builder.fill_none = true;
            }
        }
        b"tailEnd" if line_depth > 0 => builder.tail_arrow = arrow_type_present(element),
        b"headEnd" if line_depth > 0 => builder.head_arrow = arrow_type_present(element),
        _ => {}
    }
}

fn arrow_type_present(element: &BytesStart<'_>) -> bool {
    !matches!(attribute_value(element, b"type").as_deref(), Some("none"))
}

#[cfg(test)]
#[path = "docx_context_shape_tests.rs"]
mod tests;
