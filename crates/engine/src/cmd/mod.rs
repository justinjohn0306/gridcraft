//! The command registry. Ids follow Excel's ribbon (`home.bold`, `insert.chart`,
//! `data.sortAscending`) plus primitive editing commands (`cell.set`, `selection.set`).

pub mod data;
pub mod edit;
pub mod file;
pub mod format;
pub mod formulas;
pub mod inspect;
pub mod insert;
pub mod review;
pub mod sheet;
pub mod view;

use serde::Serialize;
use serde_json::Value as Json;
use sheetcraft_calc::Key;
use sheetcraft_core::{CellRef, RangeRef};
use sheetcraft_model::Workbook;

use crate::{DocState, EngineError, Result, Selection, Session};

pub type Run = fn(&mut Session, &Json) -> Result<Json>;
pub type Enabled = fn(&Session) -> std::result::Result<(), String>;

pub struct CommandSpec {
    pub id: &'static str,
    pub label: &'static str,
    /// Ribbon/menu placement, e.g. `["Home", "Font"]`. Empty = not in menus.
    pub menu: &'static [&'static str],
    pub shortcut: Option<&'static str>,
    pub params: &'static str,
    pub enabled: Enabled,
    pub run: Run,
    pub journal: bool,
    pub undoable: bool,
}

#[derive(Clone, Debug, Serialize)]
pub struct CommandInfo {
    pub id: &'static str,
    pub label: &'static str,
    pub menu: Vec<&'static str>,
    pub shortcut: Option<&'static str>,
    pub params: &'static str,
    pub enabled: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub disabled_reason: Option<String>,
}

impl CommandSpec {
    pub fn info(&self, s: &Session) -> CommandInfo {
        let e = (self.enabled)(s);
        CommandInfo {
            id: self.id,
            label: self.label,
            menu: self.menu.to_vec(),
            shortcut: self.shortcut,
            params: self.params,
            enabled: e.is_ok(),
            disabled_reason: e.err(),
        }
    }
}

pub fn always(_: &Session) -> std::result::Result<(), String> {
    Ok(())
}
pub fn has_doc(s: &Session) -> std::result::Result<(), String> {
    s.active().map(|_| ()).ok_or_else(|| "no workbook open".into())
}
pub fn can_undo(s: &Session) -> std::result::Result<(), String> {
    has_doc(s)?;
    if s.active().is_some_and(|d| !d.undo.is_empty()) { Ok(()) } else { Err("nothing to undo".into()) }
}
pub fn can_redo(s: &Session) -> std::result::Result<(), String> {
    has_doc(s)?;
    if s.active().is_some_and(|d| !d.redo.is_empty()) { Ok(()) } else { Err("nothing to redo".into()) }
}
pub fn has_clipboard(s: &Session) -> std::result::Result<(), String> {
    has_doc(s)?;
    if s.clipboard.is_some() { Ok(()) } else { Err("the clipboard is empty".into()) }
}

macro_rules! cmd {
    ($id:literal, $label:literal, [$($m:literal),*], $sc:expr, $params:literal, $en:expr, $run:expr) => {
        $crate::cmd::CommandSpec { id: $id, label: $label, menu: &[$($m),*], shortcut: $sc, params: $params, enabled: $en, run: $run, journal: true, undoable: true }
    };
    (query $id:literal, $label:literal, [$($m:literal),*], $sc:expr, $params:literal, $en:expr, $run:expr) => {
        $crate::cmd::CommandSpec { id: $id, label: $label, menu: &[$($m),*], shortcut: $sc, params: $params, enabled: $en, run: $run, journal: false, undoable: false }
    };
    (noundo $id:literal, $label:literal, [$($m:literal),*], $sc:expr, $params:literal, $en:expr, $run:expr) => {
        $crate::cmd::CommandSpec { id: $id, label: $label, menu: &[$($m),*], shortcut: $sc, params: $params, enabled: $en, run: $run, journal: true, undoable: false }
    };
}
pub(crate) use cmd;

pub fn command_specs() -> &'static [CommandSpec] {
    static SPECS: std::sync::OnceLock<Vec<CommandSpec>> = std::sync::OnceLock::new();
    SPECS.get_or_init(|| {
        let mut v = Vec::new();
        v.extend(file::specs());
        v.extend(edit::specs());
        v.extend(format::specs());
        v.extend(sheet::specs());
        v.extend(insert::specs());
        v.extend(data::specs());
        v.extend(formulas::specs());
        v.extend(review::specs());
        v.extend(view::specs());
        v.extend(inspect::specs());
        v
    })
}

pub fn find_command(id: &str) -> Option<&'static CommandSpec> {
    command_specs().iter().find(|c| c.id == id)
}

// ---------------------------------------------------------------- param helpers

pub(crate) fn bad(cmd: &str, msg: impl Into<String>) -> EngineError {
    EngineError::BadParams { cmd: cmd.into(), msg: msg.into() }
}
pub(crate) fn str_param<'a>(p: &'a Json, key: &str) -> Option<&'a str> {
    p.get(key).and_then(Json::as_str)
}
pub(crate) fn f64_param(p: &Json, key: &str) -> Option<f64> {
    p.get(key).and_then(Json::as_f64)
}
pub(crate) fn bool_param(p: &Json, key: &str) -> Option<bool> {
    p.get(key).and_then(Json::as_bool)
}
pub(crate) fn u32_param(p: &Json, key: &str) -> Option<u32> {
    p.get(key).and_then(Json::as_u64).map(|v| v.min(u32::MAX as u64) as u32)
}
pub(crate) fn ok() -> Result<Json> {
    Ok(Json::Null)
}

/// A range parameter (`"A1:B2"`, `"Sheet2!A1"`, `"B:B"`), defaulting to the current selection's
/// areas.
pub(crate) fn target_ranges(s: &Session, p: &Json) -> Result<Vec<RangeRef>> {
    if let Some(r) = str_param(p, "range").or_else(|| str_param(p, "cell")) {
        let (_, body) = split_sheet(r);
        let mut out = Vec::new();
        for part in body.split(',') {
            out.push(RangeRef::parse(part).ok_or_else(|| bad("range", format!("not a range: `{part}`")))?);
        }
        return Ok(out);
    }
    Ok(s.doc()?.selection.ranges.clone())
}

pub(crate) fn target_range(s: &Session, p: &Json) -> Result<RangeRef> {
    Ok(target_ranges(s, p)?.first().copied().unwrap_or_default())
}

/// The sheet a call names (`sheet` param or a `Sheet!` prefix on `range`), else the active one.
pub(crate) fn target_sheet(s: &Session, p: &Json) -> Result<usize> {
    let d = s.doc()?;
    if let Some(v) = p.get("sheet") {
        if let Some(i) = v.as_u64() {
            return if (i as usize) < d.wb.sheets.len() { Ok(i as usize) } else { Err(bad("sheet", "no such sheet")) };
        }
        if let Some(n) = v.as_str() {
            return d.wb.sheet_index(n).ok_or_else(|| bad("sheet", format!("no sheet named `{n}`")));
        }
    }
    if let Some(r) = str_param(p, "range").or_else(|| str_param(p, "cell"))
        && let (Some(sh), _) = split_sheet(r)
    {
        return d.wb.sheet_index(&sh).ok_or_else(|| bad("range", format!("no sheet named `{sh}`")));
    }
    Ok(d.wb.active_sheet)
}

/// `'My Sheet'!A1` → (`Some("My Sheet")`, `A1`).
pub(crate) fn split_sheet(r: &str) -> (Option<String>, &str) {
    match r.rsplit_once('!') {
        Some((sh, body)) => (Some(sh.trim_matches('\'').replace("''", "'")), body),
        None => (None, r),
    }
}

pub(crate) fn cell_param(p: &Json, key: &str) -> Option<CellRef> {
    str_param(p, key).and_then(|t| CellRef::parse(split_sheet(t).1))
}

/// Applies an edit to a copy of the active workbook, then recalculates and commits it.
/// `changed` collects cells whose content changed; set `structural` for row/column/sheet edits
/// (full recalc).
pub struct Ctx<'a> {
    pub wb: Workbook,
    pub changed: Vec<Key>,
    pub structural: bool,
    pub sel: &'a mut Selection,
}

pub(crate) fn edit<R>(s: &mut Session, f: impl FnOnce(&mut Ctx) -> Result<R>) -> Result<R> {
    let d = s.doc_mut()?;
    commit(d, f)
}

pub(crate) fn commit<R>(d: &mut DocState, f: impl FnOnce(&mut Ctx) -> Result<R>) -> Result<R> {
    let mut sel = d.selection.clone();
    let mut ctx = Ctx { wb: (*d.wb).clone(), changed: Vec::new(), structural: false, sel: &mut sel };
    let r = f(&mut ctx)?;
    let Ctx { mut wb, changed, structural, .. } = ctx;
    if structural {
        d.calc.recalc_all(&mut wb);
    } else if !changed.is_empty() {
        d.calc.cells_changed(&mut wb, &changed);
    }
    d.wb = std::sync::Arc::new(wb);
    d.selection = sel;
    Ok(r)
}

impl Ctx<'_> {
    pub fn sheet_index(&self) -> usize {
        self.wb.active_sheet
    }
    pub fn sheet_mut(&mut self, i: usize) -> Result<&mut sheetcraft_model::Sheet> {
        self.wb.sheet_mut(i).ok_or_else(|| EngineError::Other("no such sheet".into()))
    }
    pub fn touch(&mut self, sheet: usize, c: CellRef) {
        self.changed.push((sheet, c));
    }
}
