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

/// A 2×2 PNG, so objects can be injected into a workbook in tests.
fn tiny_png() -> Vec<u8> {
    let img = image::RgbaImage::from_raw(2, 2, vec![255, 0, 0, 255, 0, 255, 0, 128, 0, 0, 255, 255, 0, 0, 0, 0]).unwrap();
    let mut png = Vec::new();
    image::DynamicImage::ImageRgba8(img).write_to(&mut std::io::Cursor::new(&mut png), image::ImageFormat::Png).unwrap();
    png
}

fn add_image(s: &mut Session, cell: &str, mode: gridcraft_model::AnchorMode) -> u32 {
    let d = s.doc_mut().unwrap();
    let wb = std::sync::Arc::make_mut(&mut d.wb);
    let id = wb.next_object_id();
    wb.sheet_mut(0).unwrap().images.push(gridcraft_model::Image {
        id,
        anchor: gridcraft_model::Anchor { cell: CellRef::parse(cell).unwrap(), dx: 0.0, dy: 0.0, width: 40.0, height: 40.0, mode },
        data: tiny_png(),
        mime: "image/png".into(),
        alt: String::new(),
    });
    id
}

fn image_cell(s: &Session, id: u32) -> Option<CellRef> {
    s.doc().unwrap().wb.active().unwrap().images.iter().find(|i| i.id == id).map(|i| i.anchor.cell)
}

/// Issue #5: images anchored in cells follow their row when the sheet is sorted.
#[test]
fn sort_moves_anchored_images_with_their_rows() {
    use gridcraft_model::AnchorMode;
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [["Name", "Score"], ["b", 2], ["c", 3], ["a", 1]]})).unwrap();
    // One picture per data row, in its own row.
    let ib = add_image(&mut s, "A2", AnchorMode::MoveAndSize);
    let ic = add_image(&mut s, "A3", AnchorMode::MoveAndSize);
    let ia = add_image(&mut s, "A4", AnchorMode::MoveAndSize);
    s.execute("selection.set", json!({"range": "A1:B4"})).unwrap();
    s.execute("data.sortAscending", json!({"header": true, "column": "A"})).unwrap();
    // Ascending by name: b, c, a → a, b, c. Each picture lands in its row's new home.
    assert_eq!(image_cell(&s, ib), Some(CellRef::parse("A3").unwrap()));
    assert_eq!(image_cell(&s, ic), Some(CellRef::parse("A4").unwrap()));
    assert_eq!(image_cell(&s, ia), Some(CellRef::parse("A2").unwrap()));
}

/// An `absolute`-pinned object stays exactly where it was drawn across a sort and a row delete.
#[test]
fn absolute_objects_ignore_sort_and_structural_edits() {
    use gridcraft_model::AnchorMode;
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [["Name", "Score"], ["b", 2], ["c", 3], ["a", 1]]})).unwrap();
    let pinned = add_image(&mut s, "A3", AnchorMode::Absolute);
    s.execute("selection.set", json!({"range": "A1:B4"})).unwrap();
    s.execute("data.sortAscending", json!({"header": true, "column": "A"})).unwrap();
    assert_eq!(image_cell(&s, pinned), Some(CellRef::parse("A3").unwrap()));
    s.execute("home.deleteRows", json!({"rows": "2:2"})).unwrap();
    assert_eq!(image_cell(&s, pinned), Some(CellRef::parse("A3").unwrap()));
}

/// A "move but don't size" object follows its row across a sort but keeps its size.
#[test]
fn move_only_objects_follow_but_do_not_resize() {
    use gridcraft_model::AnchorMode;
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [["Name", "N"], ["b", 2], ["a", 1]]})).unwrap();
    let id = add_image(&mut s, "A2", AnchorMode::MoveOnly);
    s.execute("selection.set", json!({"range": "A1:B3"})).unwrap();
    s.execute("data.sortAscending", json!({"column": "A"})).unwrap();
    // "a" moved to row 2 and "b" to row 3, so the image follows from A2 to A3.
    assert_eq!(image_cell(&s, id), Some(CellRef::parse("A3").unwrap()));
    let a = s.doc().unwrap().wb.active().unwrap().images.iter().find(|i| i.id == id).unwrap().anchor;
    assert_eq!((a.width, a.height), (40.0, 40.0));
}

/// `object.setAnchorMode` changes how an object follows its cells.
#[test]
fn set_anchor_mode_changes_following() {
    use gridcraft_model::AnchorMode;
    let mut s = s();
    s.execute("range.setValues", json!({"range": "A1", "values": [["Name", "N"], ["b", 2], ["a", 1]]})).unwrap();
    let id = add_image(&mut s, "A2", AnchorMode::MoveAndSize);
    let r = s.execute("object.setAnchorMode", json!({"kind": "image", "id": id, "mode": "dontMoveOrSizeWithCells"})).unwrap();
    assert_eq!(r["mode"], "absolute");
    s.execute("selection.set", json!({"range": "A1:B3"})).unwrap();
    s.execute("data.sortAscending", json!({"column": "A"})).unwrap();
    // Pinned absolute: it stays put even though its row moved.
    assert_eq!(image_cell(&s, id), Some(CellRef::parse("A2").unwrap()));
    // A bad mode is rejected, not silently ignored.
    assert!(s.execute("object.setAnchorMode", json!({"kind": "image", "id": id, "mode": "sideways"})).is_err());
}
