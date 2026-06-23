use std::fmt::Write;

use unicode_normalization::UnicodeNormalization;

use crate::render::font_subst;

use super::*;

/// Word's default tab stop interval (0.5 inch = 36pt).
const DEFAULT_TAB_WIDTH_PT: f64 = 36.0;
const PPTX_SOFT_LINE_BREAK_CHAR: char = '\u{000B}';

pub(super) fn generate_paragraph(out: &mut String, para: &Paragraph) -> Result<(), ConvertError> {
    let style = &para.style;

    if let Some(level) = style.heading_level {
        let _ = write!(out, "#heading(level: {level})[");
        generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());
        out.push_str("]\n");
        return Ok(());
    }

    let has_para_style = needs_block_wrapper(style);
    let is_empty = para.runs.is_empty();

    if has_para_style {
        out.push_str("#block(");
        if is_empty {
            // Empty paragraphs are rendered as `#v(line_height)` inside the
            // wrapper. If we also emit the regular `below:` value
            // (zero_space_after_block_buffer_pt ~= one line) we double-count
            // and each blank line takes ~2x the vertical space Word gives
            // it. Word renders an empty paragraph at exactly one
            // line-height with `w:after="0"`; we mirror that by skipping
            // the below: term and letting `#v(...)` carry the visible
            // spacing. Inset/justify/direction/indent are still meaningful
            // (an empty paragraph can be authored with indent that affects
            // following content) and are emitted normally.
            write_block_params_empty_paragraph(out, style);
        } else {
            write_block_params(out, style);
        }
        out.push_str(")[\n");
        write_par_settings(out, style);
    }

    if is_empty {
        let _ = write!(
            out,
            "#v({}pt)",
            format_f64(empty_paragraph_height_pt(style))
        );
        if has_para_style {
            out.push_str("\n]");
        }
        out.push('\n');
        return Ok(());
    }

    let alignment = style.alignment;
    let use_align = matches!(
        alignment,
        Some(Alignment::Center) | Some(Alignment::Right) | Some(Alignment::Left)
    );

    if use_align {
        let align_str = match alignment {
            Some(Alignment::Left) => "left",
            Some(Alignment::Center) => "center",
            Some(Alignment::Right) => "right",
            _ => "left",
        };
        let _ = write!(out, "#align({align_str})[");
    }

    generate_runs_with_tabs(out, &para.runs, style.tab_stops.as_deref());

    if use_align {
        out.push(']');
    }

    if has_para_style {
        out.push_str("\n]");
    }

    out.push('\n');
    Ok(())
}

fn natural_single_line_height_pt(style: &ParagraphStyle) -> f64 {
    let font_size_pt: f64 = style.font_size.unwrap_or(12.0);
    let default_line_height_pt: f64 = font_size_pt * 1.2;

    match style.line_spacing {
        Some(LineSpacing::Exact(points)) => default_line_height_pt.max(points),
        Some(LineSpacing::Proportional(factor)) => {
            default_line_height_pt.max(font_size_pt * factor)
        }
        None => default_line_height_pt,
    }
}

fn empty_paragraph_height_pt(style: &ParagraphStyle) -> f64 {
    natural_single_line_height_pt(style)
}

fn zero_space_after_block_buffer_pt(style: &ParagraphStyle) -> f64 {
    let unclamped_buffer_pt = natural_single_line_height_pt(style).clamp(8.0, 24.0);
    (unclamped_buffer_pt * 100.0).round() / 100.0
}

fn paragraph_line_spacing_extra_pt(style: &ParagraphStyle) -> f64 {
    // The leading value Typst will apply within the paragraph, expressed in
    // points, which we also need to lay on top of #block(below:) so that
    // between-block gaps match within-paragraph baseline-to-baseline spacing.
    //
    // Empirical observation: Typst's `#block(below: X)` gives an inter-block
    // baseline-to-baseline gap of roughly natural_font_metric + X (verified
    // with a minimal `#block(below: 5pt)` vs `#block(below: 40pt)` test,
    // where the gap moved from 12.9pt to 47.9pt — a perfectly linear +35pt
    // matching the +35pt difference in `below:`). Crucially, the `leading:`
    // set via `#set par(...)` only widens spacing *within* a paragraph; it
    // does NOT contribute to inter-block spacing. Word's render makes no
    // such distinction: between-paragraph baseline-to-baseline equals
    // within-paragraph baseline-to-baseline plus w:after. To reconcile,
    // every block needs to absorb the full leading value into its below:
    // calculation, otherwise gaps after paragraphs render ~half-a-line
    // shorter than Word for 1.5x spacing.
    //
    // For Proportional(factor), write_par_settings emits
    // `leading: (factor * 0.65)em`, which is `factor * 0.65 * font_size`
    // points; we match that exactly. The previous formula
    // `font_size * (factor - 1.0)` produced half the needed value at 1.5x
    // (6pt instead of 11.7pt at 12pt), which manifested as a uniform ~6pt
    // undershoot at every block transition in the Alfresco Enterprise
    // Manual measurements.
    //
    // Single-spacing (factor <= 1.0) and the no-line-spacing case keep their
    // zero return for now; only the explicitly-spaced-out case is
    // empirically verified and present in our regression document. Exact()
    // also retains its `points - font_size` formula pending separate
    // verification.
    let font_size_pt: f64 = style.font_size.unwrap_or(12.0);
    match style.line_spacing {
        Some(LineSpacing::Proportional(factor)) if factor > 1.0 => {
            (font_size_pt * factor * 0.65).max(0.0)
        }
        Some(LineSpacing::Exact(points)) => (points - font_size_pt).max(0.0),
        _ => 0.0,
    }
}

fn effective_block_space_after_pt(style: &ParagraphStyle) -> f64 {
    // Word's <w:spacing w:after> is the *extra* gap layered on top of the
    // paragraph's natural line height: a "3pt after" paragraph with 1.5x line
    // spacing sits ~half-a-line + 3pt below its predecessor, not 3pt total.
    // Typst's #block(below:) is the *total* gap, so we add the line
    // spacing's extra contribution on top of the explicit space_after to
    // faithfully render Word's intent. This mirrors the formula already
    // used in native_list_item_spacing_pt for within-list inter-item
    // spacing, so block boundaries and list-item boundaries follow one rule.
    //
    // Negative space_after (an explicit overlap hint) is passed through
    // unchanged; the empty-paragraph buffer continues to apply when
    // space_after is unset or zero.
    let line_extras = paragraph_line_spacing_extra_pt(style);
    match style.space_after {
        Some(space_after) if space_after > 0.0 => space_after + line_extras,
        Some(space_after) if space_after < 0.0 => space_after,
        Some(_) | None => zero_space_after_block_buffer_pt(style),
    }
}

pub(super) fn needs_block_wrapper(style: &ParagraphStyle) -> bool {
    style.space_before.is_some()
        || style.space_after.is_some()
        || style.line_spacing.is_some()
        || matches!(style.alignment, Some(Alignment::Justify))
        || matches!(style.direction, Some(TextDirection::Rtl))
        || style.indent_left.is_some_and(|i| i.abs() > 0.0001)
        || style.indent_right.is_some_and(|i| i.abs() > 0.0001)
}

pub(super) fn write_block_params_empty_paragraph(out: &mut String, style: &ParagraphStyle) {
    // For empty paragraphs, the inner `#v(line_height)` filler is the
    // sole intended source of vertical space. We must therefore PIN both
    // `above:` and `below:` on the wrapper to 0pt — not merely omit them.
    // Typst falls back to its built-in default of ~1.2em per side when a
    // block parameter is absent, which silently injects ~14.4pt of extra
    // gap per empty paragraph (measured against Word: a title page with
    // 13 empty spacer paragraphs picks up roughly 75pt of accidental
    // overhead this way, pushing trailing content to the next page).
    //
    // We still emit any author-specified `above:` (w:before) on top of
    // the zero pin — Word treats `w:before` as an explicit add — by
    // letting it override the zero default.
    let mut first = true;
    write_param(out, &mut first, "width: 100%");
    let above_pt = style.space_before.unwrap_or(0.0);
    write_param(
        out,
        &mut first,
        &format!("above: {}pt", format_f64(above_pt)),
    );
    write_param(out, &mut first, "below: 0pt");
    let mut inset_parts: Vec<String> = Vec::new();
    if let Some(left) = style.indent_left {
        if left.abs() > 0.0001 {
            inset_parts.push(format!("left: {}pt", format_f64(left)));
        }
    }
    if let Some(right) = style.indent_right {
        if right.abs() > 0.0001 {
            inset_parts.push(format!("right: {}pt", format_f64(right)));
        }
    }
    if !inset_parts.is_empty() {
        write_param(
            out,
            &mut first,
            &format!("inset: ({})", inset_parts.join(", ")),
        );
    }
}

pub(super) fn write_block_params(out: &mut String, style: &ParagraphStyle) {
    let mut first = true;

    write_param(out, &mut first, "width: 100%");

    if let Some(above) = style.space_before {
        write_param(out, &mut first, &format!("above: {}pt", format_f64(above)));
    }
    if style.space_after.is_some() || style.line_spacing.is_some() {
        let below = effective_block_space_after_pt(style);
        write_param(out, &mut first, &format!("below: {}pt", format_f64(below)));
    }

    // Word's <w:ind w:left> / <w:ind w:right> on non-list paragraphs becomes
    // block inset so the paragraph indents from the left margin (and stops
    // short of the right) instead of flowing flush to the block edge. List
    // items get their indentation from the enclosing #enum/#list and don't
    // pass through this code path, so there's no double-indent risk.
    let mut inset_parts: Vec<String> = Vec::new();
    if let Some(indent_left) = style.indent_left.filter(|i| i.abs() > 0.0001) {
        inset_parts.push(format!("left: {}pt", format_f64(indent_left)));
    }
    if let Some(indent_right) = style.indent_right.filter(|i| i.abs() > 0.0001) {
        inset_parts.push(format!("right: {}pt", format_f64(indent_right)));
    }
    if !inset_parts.is_empty() {
        write_param(
            out,
            &mut first,
            &format!("inset: ({})", inset_parts.join(", ")),
        );
    }
}

pub(super) fn write_par_settings(out: &mut String, style: &ParagraphStyle) {
    if let Some(ref spacing) = style.line_spacing {
        match spacing {
            LineSpacing::Proportional(factor) => {
                let leading = factor * 0.65;
                let _ = writeln!(out, "  #set par(leading: {}em)", format_f64(leading));
            }
            LineSpacing::Exact(pts) => {
                let _ = writeln!(out, "  #set par(leading: {}pt)", format_f64(*pts));
            }
        }
    }
    if matches!(style.alignment, Some(Alignment::Justify)) {
        out.push_str("  #set par(justify: true)\n");
    }
    if matches!(style.direction, Some(TextDirection::Rtl)) {
        out.push_str("  #set text(dir: rtl)\n");
    }
}

pub(super) fn generate_runs_with_tabs(
    out: &mut String,
    runs: &[Run],
    tab_stops: Option<&[TabStop]>,
) {
    if !paragraph_contains_tabs(runs) {
        generate_runs(out, runs);
        return;
    }

    let segments: Vec<Vec<Run>> = split_runs_on_tabs(runs);
    out.push_str("#context {\n");

    for (index, segment) in segments.iter().enumerate() {
        let _ = write!(out, "  let tab_segment_{index} = [");
        generate_runs(out, segment);
        out.push_str("]\n");

        if index == 0 {
            out.push_str("  let tab_prefix_0 = tab_segment_0\n");
            continue;
        }

        write_tab_segment_bindings(out, index, segment, tab_stops);
    }

    let _ = writeln!(out, "  tab_prefix_{}", segments.len() - 1);
    out.push('}');
}

pub(super) fn generate_runs_with_tabs_no_wrap(
    out: &mut String,
    runs: &[Run],
    tab_stops: Option<&[TabStop]>,
) {
    let preserve_cjk_no_wrap: bool = runs
        .iter()
        .filter(|run| run.footnote.is_none())
        .any(|run| run.text.chars().any(is_cjk_like));
    let mut no_wrap_state: NoWrapState = NoWrapState::default();
    let transformed_runs: Vec<Run> = runs
        .iter()
        .map(|run| {
            let mut transformed_run: Run = run.clone();
            if transformed_run.footnote.is_none() {
                transformed_run.text = no_wrap_text(
                    &transformed_run.text,
                    preserve_cjk_no_wrap,
                    &mut no_wrap_state,
                );
            } else {
                no_wrap_state = NoWrapState::default();
            }
            transformed_run
        })
        .collect();

    generate_runs_with_tabs(out, &transformed_runs, tab_stops);
}

#[derive(Clone, Copy, Default)]
struct NoWrapState {
    previous_visible_char: Option<char>,
    previous_non_breaking_space: bool,
}

/// Emits Typst variable bindings for a non-first tab segment: measurement,
/// decimal anchor (if applicable), default remainder, advance, fill, and
/// the accumulated prefix content variable.
fn write_tab_segment_bindings(
    out: &mut String,
    index: usize,
    segment: &[Run],
    tab_stops: Option<&[TabStop]>,
) {
    let _ = writeln!(
        out,
        "  let tab_prefix_width_{index} = measure(tab_prefix_{}).width",
        index - 1
    );
    let _ = writeln!(
        out,
        "  let tab_segment_width_{index} = measure(tab_segment_{index}).width"
    );

    if let Some(anchor_runs) = extract_decimal_anchor_runs(segment) {
        let _ = write!(out, "  let tab_decimal_anchor_{index} = [");
        generate_runs(out, &anchor_runs);
        out.push_str("]\n");
        let _ = writeln!(
            out,
            "  let tab_decimal_width_{index} = measure(tab_decimal_anchor_{index}).width"
        );
    }

    let _ = writeln!(
        out,
        "  let tab_default_remainder_{index} = calc.rem-euclid(tab_prefix_width_{index}.abs.pt(), {})",
        format_f64(DEFAULT_TAB_WIDTH_PT)
    );
    let _ = writeln!(
        out,
        "  let tab_advance_{index} = {}",
        build_tab_advance_expr(index, segment, tab_stops)
    );
    let _ = writeln!(
        out,
        "  let tab_fill_{index} = {}",
        build_tab_fill_expr(index, tab_stops)
    );
    let _ = writeln!(
        out,
        "  let tab_prefix_{index} = [#tab_prefix_{}#tab_fill_{index}#tab_segment_{index}]",
        index - 1
    );
}

fn paragraph_contains_tabs(runs: &[Run]) -> bool {
    runs.iter().any(|run| run.text.contains('\t'))
}

pub(super) fn generate_runs(out: &mut String, runs: &[Run]) {
    for run in runs {
        generate_run(out, run);
    }
}

fn no_wrap_text(text: &str, preserve_cjk_no_wrap: bool, state: &mut NoWrapState) -> String {
    if !preserve_cjk_no_wrap {
        return text.to_string();
    }

    let mut out: String = String::new();

    for ch in text.chars() {
        if matches!(ch, '\t' | PPTX_SOFT_LINE_BREAK_CHAR) {
            out.push(ch);
            *state = NoWrapState::default();
            continue;
        }

        if ch == ' ' {
            out.push('\u{00A0}');
            state.previous_visible_char = None;
            state.previous_non_breaking_space = true;
            continue;
        }

        if state.previous_non_breaking_space
            || state
                .previous_visible_char
                .is_some_and(|prev| needs_no_wrap_joiner(prev, ch))
        {
            out.push('\u{2060}');
        }
        out.push(ch);
        state.previous_visible_char = Some(ch);
        state.previous_non_breaking_space = false;
    }

    out
}

fn needs_no_wrap_joiner(previous: char, current: char) -> bool {
    !previous.is_whitespace() && !current.is_whitespace()
}

fn is_cjk_like(ch: char) -> bool {
    matches!(
        ch as u32,
        0x1100..=0x11FF
            | 0x2E80..=0x2FFF
            | 0x3000..=0x303F
            | 0x3040..=0x30FF
            | 0x3130..=0x318F
            | 0x31F0..=0x31FF
            | 0x3400..=0x4DBF
            | 0x4E00..=0x9FFF
            | 0xAC00..=0xD7AF
            | 0xF900..=0xFAFF
            | 0xFF00..=0xFFEF
    )
}

fn split_runs_on_tabs(runs: &[Run]) -> Vec<Vec<Run>> {
    let mut segments: Vec<Vec<Run>> = vec![Vec::new()];

    for run in runs {
        if run.footnote.is_some() || !run.text.contains('\t') {
            if run.footnote.is_some() || !run.text.is_empty() {
                segments
                    .last_mut()
                    .expect("split_runs_on_tabs should always have a segment")
                    .push(run.clone());
            }
            continue;
        }

        for (index, part) in run.text.split('\t').enumerate() {
            if index > 0 {
                segments.push(Vec::new());
            }

            if !part.is_empty() {
                segments
                    .last_mut()
                    .expect("split_runs_on_tabs should always have a segment")
                    .push(Run {
                        text: part.to_string(),
                        style: run.style.clone(),
                        href: run.href.clone(),
                        footnote: None,
                    });
            }
        }
    }

    segments
}

fn extract_decimal_anchor_runs(runs: &[Run]) -> Option<Vec<Run>> {
    let visible_text: String = runs
        .iter()
        .filter(|run| run.footnote.is_none())
        .map(|run| run.text.as_str())
        .collect();
    let separator_offset: usize = find_decimal_separator_offset(&visible_text)?;

    let mut anchor_runs: Vec<Run> = Vec::new();
    let mut visible_offset: usize = 0;

    for run in runs {
        if run.footnote.is_some() {
            anchor_runs.push(run.clone());
            continue;
        }

        let run_end: usize = visible_offset + run.text.len();

        // Entire run falls before the separator — include it whole.
        if run_end <= separator_offset {
            if !run.text.is_empty() {
                anchor_runs.push(run.clone());
            }
            visible_offset = run_end;
            continue;
        }

        // This run spans the separator — include only the portion before it.
        let chars_before_separator: usize = separator_offset.saturating_sub(visible_offset);
        if chars_before_separator > 0 {
            anchor_runs.push(Run {
                text: run.text[..chars_before_separator].to_string(),
                style: run.style.clone(),
                href: run.href.clone(),
                footnote: None,
            });
        }

        return Some(anchor_runs);
    }

    None
}

fn find_decimal_separator_offset(text: &str) -> Option<usize> {
    let separator = text.char_indices().rev().find(|(offset, ch)| {
        matches!(ch, '.' | ',')
            && has_ascii_digit_before(text, *offset)
            && has_ascii_digit_after(text, *offset + ch.len_utf8())
    })?;

    if is_grouped_integer(
        &text
            .chars()
            .filter(|ch| ch.is_ascii_digit() || matches!(ch, '.' | ','))
            .collect::<String>(),
        separator.1,
    ) {
        return None;
    }

    Some(separator.0)
}

fn has_ascii_digit_before(text: &str, offset: usize) -> bool {
    text[..offset].chars().rev().any(|ch| ch.is_ascii_digit())
}

fn has_ascii_digit_after(text: &str, offset: usize) -> bool {
    text[offset..].chars().any(|ch| ch.is_ascii_digit())
}

fn is_grouped_integer(text: &str, separator: char) -> bool {
    if text
        .chars()
        .any(|ch| matches!(ch, '.' | ',') && ch != separator)
    {
        return false;
    }

    let parts: Vec<&str> = text.split(separator).collect();
    parts.len() > 1
        && parts
            .iter()
            .all(|part| !part.is_empty() && part.chars().all(|ch| ch.is_ascii_digit()))
        && parts[1..].iter().all(|part| part.len() == 3)
}

fn build_tab_advance_expr(index: usize, segment: &[Run], tab_stops: Option<&[TabStop]>) -> String {
    let prefix_width_var = format!("tab_prefix_width_{index}");
    let segment_width_var = format!("tab_segment_width_{index}");
    let decimal_width_var =
        extract_decimal_anchor_runs(segment).map(|_| format!("tab_decimal_width_{index}"));
    let default_expr = build_default_tab_advance_expr(index);

    let Some(tab_stops) = tab_stops else {
        return default_expr;
    };

    if tab_stops.is_empty() {
        return default_expr;
    }

    let mut expr = String::new();
    for (stop_index, stop) in tab_stops.iter().enumerate() {
        let branch = format!(
            "calc.max(0pt, {}pt - {prefix_width_var} - {})",
            format_f64(stop.position),
            tab_alignment_offset_expr(stop, &segment_width_var, decimal_width_var.as_deref())
        );

        if stop_index == 0 {
            let _ = write!(
                expr,
                "if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        } else {
            let _ = write!(
                expr,
                " else if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        }
    }

    let _ = write!(expr, " else {{ {default_expr} }}");
    expr
}

fn build_tab_fill_expr(index: usize, tab_stops: Option<&[TabStop]>) -> String {
    let Some(tab_stops) = tab_stops else {
        return format!("h(tab_advance_{index})");
    };

    if tab_stops.is_empty() {
        return format!("h(tab_advance_{index})");
    }

    let prefix_width_var = format!("tab_prefix_width_{index}");
    let mut expr = String::new();
    for (stop_index, stop) in tab_stops.iter().enumerate() {
        let branch = tab_fill_content_expr(index, stop.leader);

        if stop_index == 0 {
            let _ = write!(
                expr,
                "if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        } else {
            let _ = write!(
                expr,
                " else if {prefix_width_var} < {}pt {{ {branch} }}",
                format_f64(stop.position)
            );
        }
    }

    let _ = write!(expr, " else {{ h(tab_advance_{index}) }}");
    expr
}

fn tab_fill_content_expr(index: usize, leader: TabLeader) -> String {
    let leader_markup = match leader {
        TabLeader::None => return format!("h(tab_advance_{index})"),
        TabLeader::Dot => ".",
        TabLeader::Hyphen => "-",
        TabLeader::Underscore => "\\_",
    };

    format!("box(width: tab_advance_{index}, repeat[{leader_markup}])")
}

fn build_default_tab_advance_expr(index: usize) -> String {
    format!(
        "if tab_default_remainder_{index} == 0 {{ {}pt }} else {{ ({} - tab_default_remainder_{index}) * 1pt }}",
        format_f64(DEFAULT_TAB_WIDTH_PT),
        format_f64(DEFAULT_TAB_WIDTH_PT)
    )
}

fn tab_alignment_offset_expr(
    stop: &TabStop,
    segment_width_var: &str,
    decimal_width_var: Option<&str>,
) -> String {
    match stop.alignment {
        TabAlignment::Left => "0pt".to_string(),
        TabAlignment::Center => format!("{segment_width_var} / 2"),
        TabAlignment::Right => segment_width_var.to_string(),
        TabAlignment::Decimal => decimal_width_var.unwrap_or(segment_width_var).to_string(),
    }
}

pub(super) fn generate_run(out: &mut String, run: &Run) {
    if let Some(ref content) = run.footnote {
        let escaped_content = escape_typst(content);
        let _ = write!(out, "#footnote[{escaped_content}]");
        return;
    }

    if run.text.contains(PPTX_SOFT_LINE_BREAK_CHAR) {
        write_run_with_soft_line_breaks(out, run);
        return;
    }

    write_run_segment(out, run, &run.text);
}

fn write_run_with_soft_line_breaks(out: &mut String, run: &Run) {
    let mut segment_start: usize = 0;

    for (offset, ch) in run.text.char_indices() {
        if ch != PPTX_SOFT_LINE_BREAK_CHAR {
            continue;
        }

        if segment_start < offset {
            write_run_segment(out, run, &run.text[segment_start..offset]);
        }
        out.push_str("#linebreak()");
        segment_start = offset + ch.len_utf8();
    }

    if segment_start < run.text.len() {
        write_run_segment(out, run, &run.text[segment_start..]);
    }
}

fn write_run_segment(out: &mut String, run: &Run, text: &str) {
    let style = &run.style;

    let needs_all_caps: bool = matches!(style.all_caps, Some(true));
    let escaped: String = if needs_all_caps {
        escape_typst(&text.to_uppercase())
    } else {
        escape_typst(text)
    };

    let wrappers: Vec<String> = collect_formatting_wrappers(run);

    for wrapper in &wrappers {
        out.push_str(wrapper);
    }

    write_run_content(out, &escaped, style);

    for _ in &wrappers {
        out.push(']');
    }
}

/// Builds the ordered list of `#command[` openers that wrap a run's content.
/// The order matches the original nesting: link > highlight > strike >
/// underline > super/sub > smallcaps.
fn collect_formatting_wrappers(run: &Run) -> Vec<String> {
    let style: &TextStyle = &run.style;
    let mut wrappers: Vec<String> = Vec::new();

    if let Some(ref href) = run.href {
        wrappers.push(format!("#link(\"{href}\")["));
    }
    if let Some(ref highlight) = style.highlight {
        wrappers.push(format!(
            "#highlight(fill: rgb({}, {}, {}))[",
            highlight.r, highlight.g, highlight.b
        ));
    }
    if matches!(style.strikethrough, Some(true)) {
        wrappers.push("#strike[".to_string());
    }
    if matches!(style.underline, Some(true)) {
        wrappers.push("#underline[".to_string());
    }
    if matches!(style.vertical_align, Some(VerticalTextAlign::Superscript)) {
        wrappers.push("#super[".to_string());
    }
    if matches!(style.vertical_align, Some(VerticalTextAlign::Subscript)) {
        wrappers.push("#sub[".to_string());
    }
    if matches!(style.small_caps, Some(true)) {
        wrappers.push("#smallcaps[".to_string());
    }

    wrappers
}

/// Writes the innermost content of a run: either `#text(params)[escaped]`
/// when text properties are present, or the escaped text directly (with a
/// `#[...]` safety wrapper when needed to prevent Typst syntax ambiguity).
fn write_run_content(out: &mut String, escaped: &str, style: &TextStyle) {
    if has_text_properties(style) {
        out.push_str("#text(");
        write_text_params(out, style);
        out.push_str(")[");
        out.push_str(escaped);
        out.push(']');
        return;
    }

    let needs_safety_wrap: bool = !escaped.is_empty()
        && out.ends_with(']')
        && !out.ends_with("\\]")
        && matches!(escaped.as_bytes()[0], b'(' | b'.' | b'[');

    if needs_safety_wrap {
        out.push_str("#[");
        out.push_str(escaped);
        out.push(']');
    } else {
        out.push_str(escaped);
    }
}

pub(super) fn has_text_properties(style: &TextStyle) -> bool {
    matches!(style.bold, Some(true))
        || matches!(style.italic, Some(true))
        || style.font_size.is_some()
        || style.color.is_some()
        || style.font_family.is_some()
        || style.letter_spacing.is_some()
}

fn inferred_font_weight(font_family: &str) -> Option<&'static str> {
    let lower = font_family.trim().to_ascii_lowercase();
    if lower.contains("extrabold") || lower.contains("extra bold") {
        Some("extrabold")
    } else if lower.contains("semibold") || lower.contains("semi bold") {
        Some("semibold")
    } else if lower.contains("medium") {
        Some("medium")
    } else if lower.contains("light") {
        Some("light")
    } else {
        None
    }
}

fn font_weight_rank(weight: &str) -> u8 {
    match weight {
        "light" => 1,
        "medium" => 2,
        "semibold" => 3,
        "bold" => 4,
        "extrabold" => 5,
        "black" => 6,
        _ => 0,
    }
}

fn effective_font_weight(style: &TextStyle) -> Option<&'static str> {
    // Only infer weight from font family name when the font (or its alias)
    // is actually available.  When using fallback fonts, uncommonly heavy
    // weights (e.g. "extrabold" = 800) may not exist in the substitute,
    // causing Typst to fall back to its built-in serif font instead.
    let inferred = style.font_family.as_deref().and_then(|family| {
        if font_subst::is_primary_font_available(family) {
            inferred_font_weight(family)
        } else {
            None
        }
    });
    let explicit = matches!(style.bold, Some(true)).then_some("bold");
    match (explicit, inferred) {
        (Some(explicit), Some(inferred)) => {
            if font_weight_rank(explicit) >= font_weight_rank(inferred) {
                Some(explicit)
            } else {
                Some(inferred)
            }
        }
        (Some(explicit), None) => Some(explicit),
        (None, Some(inferred)) => Some(inferred),
        (None, None) => None,
    }
}

pub(super) fn write_text_params(out: &mut String, style: &TextStyle) {
    let mut first = true;

    if let Some(ref family) = style.font_family {
        let font_value = font_subst::font_with_fallbacks(family);
        write_param(out, &mut first, &format!("font: {font_value}"));
    }
    if let Some(size) = style.font_size {
        write_param(out, &mut first, &format!("size: {}pt", format_f64(size)));
    }
    if let Some(weight) = effective_font_weight(style) {
        write_param(out, &mut first, &format!("weight: \"{weight}\""));
    }
    if matches!(style.italic, Some(true)) {
        write_param(out, &mut first, "style: \"italic\"");
    }
    if let Some(ref color) = style.color {
        write_param(out, &mut first, &format_color(color));
    }
    if let Some(spacing) = style.letter_spacing {
        write_param(
            out,
            &mut first,
            &format!("tracking: {}pt", format_f64(spacing)),
        );
    }
}

pub(super) fn write_param(out: &mut String, first: &mut bool, param: &str) {
    if !*first {
        out.push_str(", ");
    }
    out.push_str(param);
    *first = false;
}

pub(super) fn format_color(color: &Color) -> String {
    format!("fill: rgb({}, {}, {})", color.r, color.g, color.b)
}

pub(super) fn format_f64(v: f64) -> String {
    if v.fract() == 0.0 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}

pub(super) fn escape_typst(text: &str) -> String {
    let normalized_text: String = text.nfc().collect();
    let mut result = String::with_capacity(normalized_text.len());
    let mut chars = normalized_text.chars().peekable();
    let mut is_first_char = true;

    while let Some(ch) = chars.next() {
        let should_escape_list_prefix: bool = is_first_char
            && matches!(ch, '-' | '+')
            && chars.peek().is_some_and(|next| next.is_whitespace());

        match ch {
            // A hard line break (`<w:br/>`, carried through the IR as '\n') must
            // force a new line. A bare newline in Typst markup collapses to a
            // space, which silently merged code lines like `echo` / `printf`
            // (issue #176).
            '\n' => result.push_str("#linebreak()"),
            '\r' => {}
            '#' | '*' | '_' | '`' | '<' | '>' | '@' | '\\' | '~' | '/' | '$' | '[' | ']' | '{'
            | '}'
                if !should_escape_list_prefix =>
            {
                result.push('\\');
                result.push(ch);
            }
            _ if should_escape_list_prefix => {
                result.push('\\');
                result.push(ch);
            }
            _ => result.push(ch),
        }

        is_first_char = false;
    }
    result
}
