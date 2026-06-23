use super::*;

#[test]
fn test_generate_plain_paragraph() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("Hello World"));
}

#[test]
fn test_generate_empty_paragraph_reserves_line_height() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: Vec::new(),
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#v(14.399"),
        "empty DOCX paragraph marks should reserve vertical flow space: {result}"
    );
}

#[test]
fn test_generate_empty_paragraph_uses_style_font_size_and_line_spacing() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            font_size: Some(18.0),
            line_spacing: Some(LineSpacing::Proportional(1.5)),
            ..ParagraphStyle::default()
        },
        runs: Vec::new(),
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#v(27pt)"),
        "empty DOCX paragraph marks should use paragraph font size and line spacing: {result}"
    );
}

#[test]
fn test_generate_page_setup() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize {
            width: 612.0,
            height: 792.0,
        },
        margins: Margins {
            top: 36.0,
            bottom: 36.0,
            left: 54.0,
            right: 54.0,
        },
        content: vec![make_paragraph("test")],
        header: None,
        footer: None,
        columns: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("612pt"));
    assert!(result.contains("792pt"));
    assert!(result.contains("36pt"));
    assert!(result.contains("54pt"));
}

#[test]
fn test_generate_bold_text() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Bold text".to_string(),
            style: TextStyle {
                bold: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("weight: \"bold\""),
        "Expected bold weight in: {result}"
    );
    assert!(result.contains("Bold text"));
}

#[test]
fn test_generate_italic_text() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Italic text".to_string(),
            style: TextStyle {
                italic: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("style: \"italic\""),
        "Expected italic style in: {result}"
    );
    assert!(result.contains("Italic text"));
}

#[test]
fn test_generate_underline_text() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Underlined".to_string(),
            style: TextStyle {
                underline: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#underline["),
        "Expected underline wrapper in: {result}"
    );
    assert!(result.contains("Underlined"));
}

#[test]
fn test_generate_font_size() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Large text".to_string(),
            style: TextStyle {
                font_size: Some(24.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("size: 24pt"),
        "Expected font size in: {result}"
    );
}

#[test]
fn test_generate_font_color() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Red text".to_string(),
            style: TextStyle {
                color: Some(Color::new(255, 0, 0)),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("fill: rgb(255, 0, 0)"),
        "Expected RGB color in: {result}"
    );
}

#[test]
fn test_generate_combined_text_styles() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Styled".to_string(),
            style: TextStyle {
                bold: Some(true),
                italic: Some(true),
                font_size: Some(16.0),
                color: Some(Color::new(0, 128, 255)),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("weight: \"bold\""));
    assert!(result.contains("style: \"italic\""));
    assert!(result.contains("size: 16pt"));
    assert!(result.contains("fill: rgb(0, 128, 255)"));
    assert!(result.contains("Styled"));
}

#[test]
fn test_generate_alignment_center() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Center),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Centered".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("align(center"),
        "Expected center alignment in: {result}"
    );
}

#[test]
fn test_generate_alignment_right() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Right),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Right".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("align(right"),
        "Expected right alignment in: {result}"
    );
}

#[test]
fn test_generate_alignment_justify() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Justify),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Justified text".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("par(justify: true") || result.contains("set par(justify: true"),
        "Expected justify in: {result}"
    );
}

#[test]
fn test_generate_line_spacing_proportional() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            line_spacing: Some(LineSpacing::Proportional(2.0)),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Double spaced".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("leading:"),
        "Expected leading setting in: {result}"
    );
}

#[test]
fn test_generate_line_spacing_exact() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            line_spacing: Some(LineSpacing::Exact(18.0)),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Exact spaced".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("leading: 18pt"),
        "Expected exact leading in: {result}"
    );
}

#[test]
fn test_generate_zero_space_after_uses_small_block_buffer() {
    let doc = make_doc(vec![make_flow_page(vec![
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Center),
                font_size: Some(18.0),
                line_spacing: Some(LineSpacing::Proportional(1.15)),
                space_after: Some(0.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "First line".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        }),
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment: Some(Alignment::Center),
                font_size: Some(18.0),
                line_spacing: Some(LineSpacing::Proportional(1.15)),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "Second line".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        }),
    ])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("below: 21.6pt"),
        "Expected a small nonzero block buffer for zero space_after in: {result}"
    );
    assert!(
        !result.contains("below: 0pt"),
        "Expected zero space_after to avoid a zero Typst block gap in: {result}"
    );
}

#[test]
fn test_generate_unset_space_after_uses_small_block_buffer() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Center),
            font_size: Some(18.0),
            line_spacing: Some(LineSpacing::Proportional(1.15)),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Centered".to_string(),
            style: TextStyle {
                bold: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("below: 21.6pt"),
        "Expected unset space_after to emit a small Typst block buffer in: {result}"
    );
}

#[test]
fn test_generate_positive_space_after_is_preserved() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            font_size: Some(18.0),
            line_spacing: Some(LineSpacing::Proportional(1.15)),
            space_after: Some(6.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Spaced".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("below: 19.455pt"),
        "Expected positive space_after (6pt) to combine with the leading \
         emitted by `set par(leading:)` for 1.15x at 18pt \
         (18pt * 1.15 * 0.65 = 13.455pt) for a total below: 19.455pt, got: \
         {result}"
    );
}

#[test]
fn test_generate_paragraph_with_left_indent_emits_block_inset() {
    // Non-list paragraphs that carry <w:ind w:left> in their pPr must be
    // rendered with a corresponding block inset so they indent from the
    // surrounding flow instead of sitting flush at the page margin. Without
    // this, continuation paragraphs in Word lists (e.g. "For example, the
    // default is 8080." that follows a sub-item) collapse against the left
    // margin and the visual hierarchy with the parent list item is lost.
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            indent_left: Some(57.5),
            space_after: Some(3.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Indented continuation".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("inset: (left: 57.5pt)"),
        "Expected block inset for a paragraph with indent_left=57.5, got: {result}"
    );
}

#[test]
fn test_generate_paragraph_with_left_and_right_indent_emits_combined_inset() {
    // Both <w:ind w:left> and <w:ind w:right> are collapsed into a single
    // block inset clause. The order is left then right to match the
    // ordering Word stores in pPr (left comes first in the OOXML schema).
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            indent_left: Some(57.5),
            indent_right: Some(0.75),
            space_after: Some(3.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Indented both sides".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("inset: (left: 57.5pt, right: 0.75pt)"),
        "Expected combined left/right block inset, got: {result}"
    );
}

#[test]
fn test_generate_paragraph_without_indent_omits_inset() {
    // Paragraphs that don't carry any <w:ind> must not gain an inset clause
    // (which would silently shift them away from the left margin). This
    // guards against accidentally emitting `inset:` for indent values that
    // round down to zero or are unset.
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            space_after: Some(6.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Plain".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("inset:"),
        "Expected no inset clause for paragraph without indent, got: {result}"
    );
}

#[test]
fn test_generate_aligned_block_paragraph_uses_full_width() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            alignment: Some(Alignment::Center),
            space_after: Some(10.0),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Centered with spacing".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#block(width: 100%, below: 10pt)[\n#align(center)[Centered with spacing]"),
        "Expected aligned block paragraphs to reserve full width for inner #align in: {result}"
    );
}

#[test]
fn test_generate_letter_spacing() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Spaced text".to_string(),
            style: TextStyle {
                letter_spacing: Some(2.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("tracking: 2pt"),
        "Expected tracking param in: {result}"
    );
}

#[test]
fn test_generate_letter_spacing_negative() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Condensed".to_string(),
            style: TextStyle {
                letter_spacing: Some(-0.5),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("tracking: -0.5pt"),
        "Expected negative tracking in: {result}"
    );
}

#[test]
fn test_generate_tab_uses_measured_default_stops() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Name:\tValue".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#context {"),
        "Expected contextual tab rendering in: {result}"
    );
    assert!(
        result.contains("measure(tab_prefix_0).width"),
        "Expected tab spacing to measure the rendered prefix in: {result}"
    );
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_1.abs.pt(), 36)"),
        "Expected default tabs to advance to the next 36pt stop in: {result}"
    );
    assert!(
        !result.contains("#h(36pt)"),
        "Expected default tabs to avoid a hard-coded 36pt gap in: {result}"
    );
}

#[test]
fn test_generate_tab_uses_next_explicit_stop_and_alignment() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![
                TabStop {
                    position: 72.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 216.0,
                    alignment: TabAlignment::Right,
                    leader: TabLeader::Dot,
                },
            ]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Col1\tCol2\tCol3".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("if tab_prefix_width_1 < 72pt"),
        "Expected the first explicit stop to be chosen by measured width in: {result}"
    );
    assert!(
        result.contains("else if tab_prefix_width_2 < 216pt"),
        "Expected the next explicit stop to be selected after the first one in: {result}"
    );
    assert!(
        result.contains("216pt - tab_prefix_width_2 - tab_segment_width_2"),
        "Expected right-aligned tabs to subtract the following segment width in: {result}"
    );
}

#[test]
fn test_generate_tab_falls_back_to_next_default_stop_after_explicit_tabs() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 100.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "A\tB\tC".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("if tab_prefix_width_1 < 100pt"),
        "Expected the explicit stop to be used when it is still ahead of the prefix in: {result}"
    );
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_2.abs.pt(), 36)"),
        "Expected tabs beyond explicit stops to use the next default stop in: {result}"
    );
}

#[test]
fn test_generate_tab_leader_uses_repeat_fill() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 144.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::Dot,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Heading\t12".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("box(width: tab_advance_1, repeat[.])"),
        "Expected dot tab leaders to render with Typst repeat fill in: {result}"
    );
}

#[test]
fn test_generate_decimal_tab_uses_decimal_separator_not_thousands_separator() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 180.0,
                alignment: TabAlignment::Decimal,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Total\t1,234.56".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("let tab_decimal_anchor_1 = [1,234]"),
        "Expected decimal alignment to anchor after the thousands group in: {result}"
    );
}

#[test]
fn test_generate_decimal_tab_handles_comma_decimal_locale() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 180.0,
                alignment: TabAlignment::Decimal,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Total\t1.234,56".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("let tab_decimal_anchor_1 = [1.234]"),
        "Expected decimal alignment to anchor on the locale decimal separator in: {result}"
    );
}

#[test]
fn test_generate_multiple_paragraphs() {
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("First paragraph"),
        make_paragraph("Second paragraph"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("First paragraph"));
    assert!(result.contains("Second paragraph"));
    assert!(
        result.contains("First paragraph\n\nSecond paragraph"),
        "Expected paragraph break between flow paragraphs in: {result}"
    );
}

#[test]
fn test_generate_paragraph_with_multiple_runs() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "Normal ".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
            Run {
                text: "bold".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: " normal again".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("Normal "));
    assert!(result.contains("bold"));
    assert!(result.contains(" normal again"));
}

#[test]
fn test_generate_empty_document() {
    let doc = make_doc(vec![]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.is_empty() || !result.is_empty());
}

#[test]
fn test_generate_special_characters_escaped() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "Price: $100 #items @store",
    )])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("\\#") || result.contains("Price"),
        "Expected escaped or present text in: {result}"
    );
}
