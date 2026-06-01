use super::*;
use crate::ir::{ArrowHead, ShapeKind};

/// A document.xml body wrapper around `inner` run/drawing markup.
fn body(inner: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body><w:p><w:r>{inner}</w:r></w:p></w:body></w:document>"#
    )
}

/// A filled rectangle authored by LibreOffice (issue #176, "Shape 2").
const RECT_DRAWING: &str = r#"<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:anchor>
<wp:positionH relativeFrom="column"><wp:posOffset>2886710</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>141605</wp:posOffset></wp:positionV>
<wp:extent cx="1590675" cy="733425"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="1590840" cy="733320"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="729fcf"/></a:solidFill>
<a:ln w="0"><a:solidFill><a:srgbClr val="3465a4"/></a:solidFill></a:ln>
</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>
<mc:Fallback><w:pict><v:rect/></w:pict></mc:Fallback></mc:AlternateContent>"#;

/// A horizontal connector with a triangular arrowhead (issue #176, "Horizontal line 1").
const LINE_DRAWING: &str = r#"<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:anchor>
<wp:positionH relativeFrom="column"><wp:posOffset>1957070</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>-74295</wp:posOffset></wp:positionV>
<wp:extent cx="929640" cy="0"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="929520" cy="0"/></a:xfrm>
<a:prstGeom prst="line"><a:avLst/></a:prstGeom>
<a:ln w="0"><a:solidFill><a:srgbClr val="3465a4"/></a:solidFill>
<a:tailEnd len="med" type="triangle" w="med"/></a:ln>
</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>
<mc:Fallback><w:pict><v:line/></w:pict></mc:Fallback></mc:AlternateContent>"#;

/// A text-box shape (`wps:txbx`) — handled by docx-rs, must be ignored here.
const TEXTBOX_DRAWING: &str = r#"<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:anchor>
<wp:positionH relativeFrom="column"><wp:posOffset>2985770</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>-25400</wp:posOffset></wp:positionV>
<wp:extent cx="1390650" cy="485775"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:spPr>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="0"><a:noFill/></a:ln>
</wps:spPr>
<wps:txbx><w:txbxContent><w:p><w:r><w:t>Very important text inside a box</w:t></w:r></w:p></w:txbxContent></wps:txbx>
</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent>"#;

/// An inline picture — handled by docx-rs, must be ignored here.
const PIC_DRAWING: &str = r#"<w:drawing><wp:anchor>
<wp:positionH relativeFrom="column"><wp:posOffset>60325</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>635</wp:posOffset></wp:positionV>
<wp:extent cx="3432175" cy="2574290"/>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic><pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
</a:graphicData></a:graphic></wp:anchor></w:drawing>"#;

fn approx(actual: f64, expected: f64) {
    assert!(
        (actual - expected).abs() < 0.05,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn scans_filled_rectangle_geometry_position_and_colors() {
    let shapes = scan_drawing_shapes(&body(RECT_DRAWING));
    assert_eq!(shapes.len(), 1, "expected one rectangle shape");

    let shape = &shapes[0];
    assert!(matches!(shape.shape.kind, ShapeKind::Rectangle));
    approx(shape.offset_x, 227.30); // 2886710 EMU
    approx(shape.offset_y, 11.15); // 141605 EMU
    approx(shape.width, 125.25); // 1590675 EMU
    approx(shape.height, 57.75); // 733425 EMU

    let fill = shape.shape.fill.expect("rectangle should have a fill");
    assert_eq!((fill.r, fill.g, fill.b), (0x72, 0x9f, 0xcf));

    let stroke = shape
        .shape
        .stroke
        .as_ref()
        .expect("rectangle has an outline");
    assert_eq!(
        (stroke.color.r, stroke.color.g, stroke.color.b),
        (0x34, 0x65, 0xa4)
    );
    assert!(stroke.width > 0.0, "w=0 must map to a visible hairline");
}

#[test]
fn scans_line_with_tail_arrowhead() {
    let shapes = scan_drawing_shapes(&body(LINE_DRAWING));
    assert_eq!(shapes.len(), 1, "expected one line shape");

    let shape = &shapes[0];
    match shape.shape.kind {
        ShapeKind::Line {
            head_end, tail_end, ..
        } => {
            assert_eq!(tail_end, ArrowHead::Triangle, "tailEnd triangle → arrow");
            assert_eq!(head_end, ArrowHead::None);
        }
        ref other => panic!("expected a line, got {other:?}"),
    }
    assert!(shape.shape.fill.is_none(), "a line has no fill");
    let stroke = shape.shape.stroke.as_ref().expect("line needs a stroke");
    assert_eq!(
        (stroke.color.r, stroke.color.g, stroke.color.b),
        (0x34, 0x65, 0xa4)
    );
}

#[test]
fn ignores_text_box_and_picture_drawings() {
    // Text boxes and pictures are handled by docx-rs; this side-channel must
    // not double-emit them.
    assert!(scan_drawing_shapes(&body(TEXTBOX_DRAWING)).is_empty());
    assert!(scan_drawing_shapes(&body(PIC_DRAWING)).is_empty());
}

#[test]
fn scans_multiple_shapes_in_document_order() {
    let combined = format!("{RECT_DRAWING}{TEXTBOX_DRAWING}{LINE_DRAWING}");
    let shapes = scan_drawing_shapes(&body(&combined));
    // Only the two geometry-only shapes survive, in order: rect then line.
    assert_eq!(shapes.len(), 2);
    assert!(matches!(shapes[0].shape.kind, ShapeKind::Rectangle));
    assert!(matches!(shapes[1].shape.kind, ShapeKind::Line { .. }));
}

#[test]
fn consume_next_yields_shapes_then_none() {
    let ctx = DrawingShapeContext::from_xml(Some(&body(RECT_DRAWING)));
    assert!(ctx.consume_next().is_some());
    assert!(
        ctx.consume_next().is_none(),
        "cursor past the end yields None"
    );
}

#[test]
fn empty_when_no_drawings() {
    assert!(scan_drawing_shapes(&body("<w:t>plain text</w:t>")).is_empty());
}
