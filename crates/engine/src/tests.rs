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
fn rich_clipboard_copies_selected_displayed_cells_and_styles() {
    let mut s = s();
    s.execute("cell.set", json!({"cell": "A1", "input": "outside selection"})).unwrap();
    s.execute("range.setValues", json!({"range": "B2", "values": [[0.125, "<tag>&\"\nnext"], [0.25, "last"]]})).unwrap();
    s.execute("home.numberFormat", json!({"range": "B2:B3", "code": "0.0%"})).unwrap();
    s.execute("home.bold", json!({"range": "B2", "on": true})).unwrap();
    s.execute("home.italic", json!({"range": "B2", "on": true})).unwrap();
    s.execute("home.fontColor", json!({"range": "B2", "color": "#123456"})).unwrap();
    s.execute("home.fillColor", json!({"range": "B2", "color": "#FEDCBA"})).unwrap();
    let copied = s.execute("edit.copy", json!({"range": "B2:C3"})).unwrap();
    let html = copied["html"].as_str().expect("copy publishes HTML as well as plain text");
    assert_eq!(copied["text"], "12.5%\t\"<tag>&\"\"\nnext\"\n25.0%\tlast\n");
    assert!(html.starts_with("<table "));
    assert!(html.ends_with("</table>"));
    assert_eq!(html.matches("<tr>").count(), 2);
    assert_eq!(html.matches("<td ").count(), 4);
    assert!(html.contains("&lt;tag&gt;&amp;&quot;<br>next"));
    for content in ["12.5%", "25.0%", "font-weight:bold;", "font-style:italic;", "color:#123456;", "background:#FEDCBA;", "text-align:right;"] {
        assert!(html.contains(content), "missing {content}");
    }
    assert!(!html.contains("outside selection"));
    assert!(!html.contains("<html"));

    let cut = s.execute("edit.cut", json!({"range": "B2:C3"})).unwrap();
    assert_eq!(cut["html"], copied["html"]);
    assert_eq!(cut["text"], copied["text"]);
    s.execute("edit.paste", json!({"at": "E5", "text": cut["text"]})).unwrap();
    assert_eq!(v(&s, "E5"), Value::Number(0.125));
    assert_eq!(v(&s, "B2"), Value::Empty);
}

#[test]
fn rich_clipboard_clips_merges_and_skips_hidden_rows() {
    let mut s = s();
    s.execute("cell.set", json!({"cell": "B2", "input": "merged"})).unwrap();
    s.execute("home.mergeCenter", json!({"range": "B2:D5"})).unwrap();
    s.execute("home.hideRows", json!({"rows": "3:3"})).unwrap();
    s.execute("cell.set", json!({"cell": "E4", "input": "side"})).unwrap();
    let copied = s.execute("edit.copy", json!({"range": "B2:C4"})).unwrap();
    let html = copied["html"].as_str().expect("merged copy publishes HTML");
    assert_eq!(copied["text"], "merged\t\n\t\n");
    assert_eq!(html.matches("<tr>").count(), 2);
    assert_eq!(html.matches("<td ").count(), 1);
    assert!(html.contains("rowspan=\"2\" colspan=\"2\""));
    assert!(html.contains(">merged</td>"));

    // A selection beginning inside a merge must keep its shape without copying
    // the original anchor's value from outside the selected rectangle.
    let copied = s.execute("edit.copy", json!({"range": "C3:E5"})).unwrap();
    let html = copied["html"].as_str().expect("partially selected merge publishes HTML");
    assert_eq!(html.matches("<tr>").count(), 2);
    assert_eq!(html.matches("<td ").count(), 3);
    assert!(html.contains("rowspan=\"2\" colspan=\"2\""));
    assert!(!html.contains("merged"));
    assert!(html.contains(">side</td>"));
}

#[test]
fn rich_clipboard_omits_html_for_large_ranges_without_truncating_text() {
    let mut s = s();
    s.execute("cell.set", json!({"cell": "A10001", "input": "last cell"})).unwrap();
    let copied = s.execute("edit.copy", json!({"range": "A1:A10001"})).unwrap();
    assert!(copied.get("html").is_none(), "oversized HTML must be omitted, not partially copied");
    let text = copied["text"].as_str().unwrap();
    assert_eq!(text.lines().count(), 10_001);
    assert!(text.ends_with("last cell\n"));
    let copied = s.execute("edit.copy", json!({"range": "A10001"})).unwrap();
    assert!(copied["html"].as_str().is_some_and(|html| html.contains("last cell")));
}

#[test]
fn rich_clipboard_omits_html_when_escaping_exceeds_byte_budget() {
    let mut s = s();
    let text = "&".repeat(1_000_000);
    s.execute("range.setValues", json!({"range": "A1", "values": [[text]]})).unwrap();
    let copied = s.execute("edit.copy", json!({"range": "A1"})).unwrap();
    assert!(copied.get("html").is_none(), "an oversized escaped value must not produce partial HTML");
    assert_eq!(copied["text"].as_str().unwrap().len(), 1_000_001);
    assert!(copied["text"].as_str().unwrap().ends_with("&\n"));
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
