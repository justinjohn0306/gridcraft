use gridcraft_core::{CellRef, Value};
use serde_json::json;

use crate::Session;

fn s() -> Session {
    let mut s = Session::new();
    s.new_workbook();
    s
}

fn v(s: &Session, a: &str) -> Value {
    s.doc().unwrap().wb.active().unwrap().value(CellRef::parse(a).unwrap())
}

#[test]
fn enter_values_and_formulas() {
    let mut s = s();
    s.execute("cell.set", json!({"cell": "A1", "input": "10"})).unwrap();
    s.execute("cell.set", json!({"cell": "A2", "input": "32"})).unwrap();
    s.execute("cell.set", json!({"cell": "A3", "input": "=sum(a1:a2"})).unwrap();
    assert_eq!(v(&s, "A3"), Value::Number(42.0));
    assert!(s.execute("cell.set", json!({"cell": "A4", "input": "=1+*2"})).is_err());
    s.execute("cell.set", json!({"cell": "B1", "input": "12%"})).unwrap();
    let d = s.doc().unwrap();
    let sh = d.wb.active().unwrap();
    assert_eq!(crate::display::cell_text(&d.wb, sh, CellRef::parse("B1").unwrap()), "12%");
}

#[test]
fn undo_redo() {
    let mut s = s();
    s.execute("cell.set", json!({"cell": "A1", "input": "1"})).unwrap();
    s.execute("cell.set", json!({"cell": "A1", "input": "2"})).unwrap();
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(v(&s, "A1"), Value::Number(1.0));
    s.execute("edit.redo", json!({})).unwrap();
    assert_eq!(v(&s, "A1"), Value::Number(2.0));
    assert!(s.doc().unwrap().is_dirty());
}

#[test]
fn copy_paste_shifts_formulas() {
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [[1, 2], [3, 4]]})).unwrap();
    s.execute("cell.set", json!({"cell": "C1", "input": "=A1+B1"})).unwrap();
    s.execute("edit.copy", json!({"range": "C1"})).unwrap();
    s.execute("selection.set", json!({"range": "C2"})).unwrap();
    s.execute("edit.paste", json!({})).unwrap();
    assert_eq!(v(&s, "C2"), Value::Number(7.0));
    s.execute("edit.cut", json!({"range": "A1"})).unwrap();
    s.execute("selection.set", json!({"range": "E5"})).unwrap();
    s.execute("edit.paste", json!({})).unwrap();
    assert_eq!(v(&s, "E5"), Value::Number(1.0));
    let f = s.execute("cell.get", json!({"cell": "C1"})).unwrap();
    assert_eq!(f["formula"], "=E5+B1");
}

#[test]
fn insert_delete_rows_adjust() {
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [[1], [2], [3]]})).unwrap();
    s.execute("cell.set", json!({"cell": "B1", "input": "=SUM(A1:A3)"})).unwrap();
    s.execute("home.insertRows", json!({"rows": "2:2"})).unwrap();
    assert_eq!(s.execute("cell.get", json!({"cell": "B1"})).unwrap()["formula"], "=SUM(A1:A4)");
    s.execute("cell.set", json!({"cell": "A2", "input": "10"})).unwrap();
    assert_eq!(v(&s, "B1"), Value::Number(16.0));
    s.execute("home.deleteRows", json!({"rows": "1:1"})).unwrap();
    assert_eq!(v(&s, "B1"), Value::Empty);
    assert_eq!(v(&s, "A1"), Value::Number(10.0));
}

#[test]
fn inserting_or_deleting_cells_cancels_copy_mode() {
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [[1], [2], [3]]})).unwrap();
    s.execute("edit.cut", json!({"range": "3:3"})).unwrap();
    s.execute("home.insertRows", json!({"rows": "3:3"})).unwrap();
    assert!(s.clipboard.is_none());
    // The cut data is now in row 4; a paste must not move the new blank row 3 instead.
    assert!(s.execute("edit.paste", json!({"at": "A10"})).is_err());
    assert_eq!(v(&s, "A4"), Value::Number(3.0));
    s.execute("edit.copy", json!({"range": "A:A"})).unwrap();
    s.execute("home.deleteColumns", json!({"cols": "B:B"})).unwrap();
    assert!(s.clipboard.is_none());
    s.execute("edit.copy", json!({"range": "A1"})).unwrap();
    s.execute("home.insertCells", json!({"range": "A1", "shift": "down"})).unwrap();
    assert!(s.clipboard.is_none());
    s.execute("edit.copy", json!({"range": "A2"})).unwrap();
    s.execute("home.deleteCells", json!({"range": "A1", "shift": "up"})).unwrap();
    assert!(s.clipboard.is_none());
}

#[test]
fn fill_series() {
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [[1], [3]]})).unwrap();
    s.execute("edit.autoFill", json!({"source": "A1:A2", "target": "A1:A6"})).unwrap();
    assert_eq!(v(&s, "A6"), Value::Number(11.0));
    s.execute("cell.set", json!({"cell": "B1", "input": "Mon"})).unwrap();
    s.execute("edit.autoFill", json!({"source": "B1", "target": "B1:B3"})).unwrap();
    assert_eq!(v(&s, "B3"), Value::from("Wed"));
    s.execute("cell.set", json!({"cell": "C1", "input": "Item 9"})).unwrap();
    s.execute("edit.autoFill", json!({"source": "C1", "target": "C1:C2"})).unwrap();
    assert_eq!(v(&s, "C2"), Value::from("Item 10"));
}

#[test]
fn insert_delete_cells_adjust() {
    let formula = |s: &mut Session, a: &str| s.execute("cell.get", json!({"cell": a})).unwrap()["formula"].clone();
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [[1], [2], [3], [4], [5]]})).unwrap();
    s.execute("cell.set", json!({"cell": "C1", "input": "=SUM(A1:A5)"})).unwrap();
    s.execute("cell.set", json!({"cell": "C2", "input": "=A4*10"})).unwrap();
    s.execute("cell.set", json!({"cell": "C3", "input": "=SUM(A1:B5)"})).unwrap();
    s.execute("formulas.defineName", json!({"name": "Fourth", "refersTo": "=Sheet1!$A$4"})).unwrap();
    s.execute(
        "home.conditionalFormat",
        json!({"range": "A4:A5", "rule": {"type": "cellIs", "operator": "greater", "value": "3", "preset": "redText"}}),
    )
    .unwrap();
    s.execute("data.validation", json!({"range": "A5", "type": "whole", "operator": "greater", "formula1": "0"})).unwrap();
    s.execute("insert.chart", json!({"range": "A1:A5", "type": "line"})).unwrap();
    // Insert A3 shifting down: ranges across the insertion grow, references below move.
    s.execute("home.insertCells", json!({"range": "A3", "shift": "down"})).unwrap();
    assert_eq!(formula(&mut s, "C1"), "=SUM(A1:A6)");
    assert_eq!(formula(&mut s, "C2"), "=A5*10");
    assert_eq!(formula(&mut s, "C3"), "=SUM(A1:B5)");
    assert_eq!(v(&s, "C2"), Value::Number(40.0));
    let sh = s.doc().unwrap().wb.active().unwrap().clone();
    assert_eq!(s.doc().unwrap().wb.names[0].formula, "Sheet1!$A$5");
    assert_eq!(sh.cond_formats[0].ranges[0].a1(), "A5:A6");
    assert_eq!(sh.validations[0].ranges[0].a1(), "A6");
    assert_eq!(sh.charts[0].series[0].values, "Sheet1!$A$1:$A$6");
    // Delete A5 (the 4) shifting up: references to it become #REF!, those below move up.
    s.execute("home.deleteCells", json!({"range": "A5", "shift": "up"})).unwrap();
    assert_eq!(formula(&mut s, "C1"), "=SUM(A1:A5)");
    assert_eq!(formula(&mut s, "C2"), "=#REF!*10");
    assert_eq!(s.doc().unwrap().wb.names[0].formula, "#REF!");
    assert_eq!(v(&s, "C1"), Value::Number(11.0));
    let sh = s.doc().unwrap().wb.active().unwrap().clone();
    assert_eq!(sh.cond_formats[0].ranges[0].a1(), "A5");
    assert_eq!(sh.validations[0].ranges[0].a1(), "A5");
    assert_eq!(sh.charts[0].series[0].values, "Sheet1!$A$1:$A$5");
    // Shifting right and left.
    s.execute("range.setValues", json!({"range": "E1", "values": [[1, 2, 3]]})).unwrap();
    s.execute("cell.set", json!({"cell": "E3", "input": "=G1+SUM(E1:G1)"})).unwrap();
    s.execute("home.insertCells", json!({"range": "F1", "shift": "right"})).unwrap();
    assert_eq!(formula(&mut s, "E3"), "=H1+SUM(E1:H1)");
    s.execute("home.deleteCells", json!({"range": "E1:F1", "shift": "left"})).unwrap();
    assert_eq!(formula(&mut s, "E3"), "=F1+SUM(E1:F1)");
    assert_eq!(v(&s, "E3"), Value::Number(8.0));
}

#[test]
fn sort_and_filter() {
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [["Name", "Score"], ["b", 2], ["c", 3], ["a", 1]]})).unwrap();
    s.execute("selection.set", json!({"cell": "B2"})).unwrap();
    s.execute("data.sortDescending", json!({})).unwrap();
    assert_eq!(v(&s, "A2"), Value::from("c"));
    assert_eq!(v(&s, "A1"), Value::from("Name"));
    s.execute("data.filter", json!({})).unwrap();
    let r = s.execute("data.filterBy", json!({"column": "B", "custom": {"op": ">", "value": "1"}})).unwrap();
    assert_eq!(r["hiddenRows"], 1);
}

#[test]
fn formatting_and_styles() {
    let mut s = s();
    s.execute("selection.set", json!({"range": "A1:B2"})).unwrap();
    s.execute("home.bold", json!({})).unwrap();
    s.execute("home.fillColor", json!({"color": "#FFFF00"})).unwrap();
    s.execute("home.mergeCenter", json!({})).unwrap();
    let c = s.execute("cell.get", json!({"cell": "B2"})).unwrap();
    assert_eq!(c["style"]["font"]["bold"], true);
    assert_eq!(c["merge"], "A1:B2");
    s.execute("home.cellStyle", json!({"range": "D1", "name": "Good"})).unwrap();
}

#[test]
fn tables_and_structured_refs() {
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [["Item", "Qty"], ["x", 2], ["y", 5]]})).unwrap();
    s.execute("insert.table", json!({"range": "A1:B3"})).unwrap();
    s.execute("cell.set", json!({"cell": "D1", "input": "=SUM(Table1[Qty])"})).unwrap();
    assert_eq!(v(&s, "D1"), Value::Number(7.0));
    s.execute("table.totalRow", json!({"on": true, "table": "Table1"})).unwrap();
    assert_eq!(v(&s, "B4"), Value::Number(7.0));
}

#[test]
fn sheets() {
    let mut s = s();
    s.execute("home.insertSheet", json!({})).unwrap();
    s.execute("cell.set", json!({"cell": "A1", "input": "5"})).unwrap();
    s.execute("sheet.rename", json!({"name": "Data"})).unwrap();
    s.execute("sheet.activate", json!({"sheet": 0})).unwrap();
    s.execute("cell.set", json!({"cell": "A1", "input": "=Data!A1*2"})).unwrap();
    assert_eq!(v(&s, "A1"), Value::Number(10.0));
    s.execute("sheet.rename", json!({"sheet": "Data", "name": "My Data"})).unwrap();
    assert_eq!(s.execute("cell.get", json!({"cell": "A1"})).unwrap()["formula"], "='My Data'!A1*2");
}

#[test]
fn xlsx_roundtrip_through_engine() {
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [["a", 1], ["b", 2]]})).unwrap();
    s.execute("cell.set", json!({"cell": "B3", "input": "=SUM(B1:B2)"})).unwrap();
    let r = s.execute("file.saveBytes", json!({"format": "xlsx"})).unwrap();
    let b64 = r["base64"].as_str().unwrap().to_string();
    s.execute("file.open", json!({"name": "x.xlsx", "base64": b64})).unwrap();
    assert_eq!(v(&s, "B3"), Value::Number(3.0));
}

#[test]
fn samples_build() {
    for (name, _) in crate::sample::SAMPLES {
        let wb = crate::sample::build(name).unwrap_or_else(|| panic!("sample {name}"));
        assert!(!wb.sheets[0].cells.is_empty());
    }
}

#[test]
fn every_command_survives_empty_params() {
    let mut s = Session::new();
    s.execute("file.new", json!({"sample": "sales"})).unwrap();
    for spec in crate::command_specs() {
        if matches!(spec.id, "file.close" | "file.open" | "file.save" | "file.saveAs" | "file.exportCsv" | "file.exportHtml" | "insert.picture") {
            continue;
        }
        let _ = s.execute(spec.id, json!({}));
        let _ = s.execute(spec.id, json!([1, 2]));
        let _ = s.execute(spec.id, json!({"range": "ZZZ999999999", "cell": "", "sheet": 99}));
    }
}

#[test]
fn parity_counts() {
    let (done, total) = crate::catalog::parity();
    assert!(total > 250);
    assert!(done > 150, "parity {done}/{total}");
}

#[test]
fn ink_strokes_and_ink_to_shape() {
    let mut s = s();
    // A rough closed box.
    let pts: Vec<[f64; 2]> = vec![[100.0, 100.0], [200.0, 102.0], [201.0, 180.0], [99.0, 181.0], [101.0, 104.0]];
    s.execute("draw.stroke", json!({"points": pts})).unwrap();
    let sh = s.doc().unwrap().wb.active().unwrap().clone();
    assert_eq!(sh.shapes.len(), 1);
    assert_eq!(sh.shapes[0].kind, gridcraft_model::ShapeKind::Ink);
    let r = s.execute("draw.inkToShape", json!({})).unwrap();
    assert_eq!(r["kind"], "Rectangle");
    // An open stroke becomes a line.
    s.execute("draw.stroke", json!({"points": [[0.0, 0.0], [50.0, 10.0], [120.0, 30.0]]})).unwrap();
    assert_eq!(s.execute("draw.inkToShape", json!({})).unwrap()["kind"], "Line");
    assert!(s.execute("draw.stroke", json!({"points": [[1.0, 1.0]]})).is_err());
}
