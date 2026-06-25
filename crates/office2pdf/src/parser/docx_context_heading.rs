//! Tracks document-wide multilevel heading counter state.
//!
//! Word's heading styles (Heading1-9 by default) typically reference a
//! multilevel numbering scheme via <w:numPr> in styles.xml — for example
//! `numId=1` with `<w:lvlText>%1.%2.%3.%4</w:lvlText>` at ilvl=3. Heading
//! paragraphs in the document body don't carry an inline <w:numPr>; they
//! inherit it from the style. To match Word's rendered output (e.g.
//! "2.2.4.12 Supported platforms") the converter must maintain counter
//! state across the whole document, increment the appropriate level when
//! each heading appears, reset all deeper levels, and format the
//! resulting counter using the level's lvlText pattern.
//!
//! This module owns that state. It exposes a single `advance_and_format`
//! entry point that takes a `(num_id, level)` and the abstract numbering
//! definition, and returns the rendered prefix string (without trailing
//! space). Levels are tracked per `num_id` so multiple multilevel schemes
//! coexist (rare but legal in OOXML).

use std::cell::RefCell;
use std::collections::HashMap;

/// State for one multilevel numbering scheme: one decimal counter per
/// possible ilvl (0..=8). When a level n is advanced, counters at depths
/// > n are reset to their `start` values on the next reference, which we
/// implement by zeroing them here and letting `lvl_text_lookup` supply
/// the right `start`.
#[derive(Debug, Clone)]
struct PerSchemeState {
    /// Current counter value at each ilvl. 0 means "not yet started"; the
    /// first advance at this level sets it to the level's `start`.
    counters: [u32; 9],
    /// Whether each level has been touched at least once. Used to decide
    /// between "first advance -> use start value" vs "subsequent advance
    /// -> increment".
    initialized: [bool; 9],
}

impl Default for PerSchemeState {
    fn default() -> Self {
        Self {
            counters: [0; 9],
            initialized: [false; 9],
        }
    }
}

pub(in super::super) struct HeadingCounterContext {
    states: RefCell<HashMap<usize, PerSchemeState>>,
}

impl HeadingCounterContext {
    pub(in super::super) fn new() -> Self {
        Self {
            states: RefCell::new(HashMap::new()),
        }
    }

    /// Advance the counter at `(num_id, level)`, returning a Vec<u32> of
    /// counters from level 0 through `level` inclusive. The caller is
    /// expected to format these into the lvlText pattern (e.g. "%1.%2").
    ///
    /// On advance: level `level` increments by 1 (or starts at `start_at`
    /// if it's its first appearance); levels deeper than `level` reset
    /// (initialized=false, counter=0) so their next first-touch picks up
    /// the right start.
    pub(in super::super) fn advance(&self, num_id: usize, level: u32, start_at: u32) -> Vec<u32> {
        let level = level as usize;
        if level > 8 {
            return Vec::new();
        }
        let mut states = self.states.borrow_mut();
        let state = states.entry(num_id).or_default();
        if state.initialized[level] {
            state.counters[level] = state.counters[level].saturating_add(1);
        } else {
            state.counters[level] = start_at.max(1);
            state.initialized[level] = true;
        }
        // Reset deeper levels.
        for deeper in (level + 1)..9 {
            state.counters[deeper] = 0;
            state.initialized[deeper] = false;
        }
        (0..=level).map(|i| state.counters[i]).collect()
    }
}
