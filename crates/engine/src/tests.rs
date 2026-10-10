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

/// Original minimal BIFF12 records, built in source rather than a vendor workbook fixture.
fn xlsb_fixture() -> Vec<u8> {
    use std::io::{Cursor, Write};
    fn wide(s: &str) -> Vec<u8> {
        let units: Vec<_> = s.encode_utf16().collect();
        (units.len() as u32).to_le_bytes().into_iter().chain(units.into_iter().flat_map(u16::to_le_bytes)).collect()
    }
    fn record(out: &mut Vec<u8>, id: u16, payload: &[u8]) {
        if id < 128 {
            out.push(id as u8);
        } else {
            out.extend([(id as u8 & 127) | 128, (id >> 7) as u8]);
        }
        let mut size = payload.len();
        while size >= 128 {
            out.push((size as u8 & 127) | 128);
            size >>= 7;
        }
        out.push(size as u8);
        out.extend(payload);
    }
    let mut book = vec![];
    record(&mut book, 131, &[]); // BeginBook
    record(&mut book, 143, &[]); // BeginBundleShs
    let mut bundle = vec![0; 4]; // Visible
    bundle.extend(1u32.to_le_bytes());
    bundle.extend(wide("rId1"));
    bundle.extend(wide("Imported"));
    record(&mut book, 156, &bundle);
    record(&mut book, 144, &[]);
    record(&mut book, 132, &[]);
    let mut sheet = vec![];
    record(&mut sheet, 129, &[]); // BeginSheet
    record(&mut sheet, 145, &[]); // BeginSheetData
    record(&mut sheet, 0, &[0; 17]); // Row 0, no spans
    let mut numeric = vec![0; 8]; // A1, default style
    numeric.extend(42.0f64.to_le_bytes());
    numeric.extend([0; 2]); // Formula flags
    numeric.extend(3u32.to_le_bytes());
    numeric.extend([0x1e, 42, 0]); // PtgInt(42)
    numeric.extend(0u32.to_le_bytes());
    record(&mut sheet, 9, &numeric); // FmlaNum
    let mut text = 1u32.to_le_bytes().to_vec(); // B1
    text.extend(0u32.to_le_bytes());
    text.extend(wide("Original data"));
    record(&mut sheet, 6, &text);
    record(&mut sheet, 146, &[]);
    record(&mut sheet, 130, &[]);
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/><Override PartName="/xl/worksheets/sheet1.bin" ContentType="application/vnd.ms-excel.worksheet"/></Types>"#;
    let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>"#;
    let book_rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/></Relationships>"#;
    for (name, data) in [
        ("[Content_Types].xml", types.as_slice()),
        ("_rels/.rels", rels.as_slice()),
        ("xl/_rels/workbook.bin.rels", book_rels.as_slice()),
        ("xl/workbook.bin", book.as_slice()),
        ("xl/worksheets/sheet1.bin", sheet.as_slice()),
    ] {
        zip.start_file(name, zip::write::SimpleFileOptions::default()).unwrap();
        zip.write_all(data).unwrap();
    }
    zip.finish().unwrap().into_inner()
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

#[test]
fn xlsb_import_keeps_cached_values_and_reports_limits() {
    let bytes = xlsb_fixture();
    assert_eq!(gridcraft_xlsx::sniff(&bytes), gridcraft_xlsx::Format::Xlsb);
    for name in ["source.xlsb", "source.xlsx"] {
        let mut s = s();
        let r = s.execute("file.open", json!({"name": name, "base64": crate::io::base64_encode(&bytes)})).unwrap();
        assert!(!r["warnings"].as_array().unwrap().is_empty());
        assert!(s.take_ui_requests().iter().any(|r| matches!(r, crate::UiRequest::Message(_))));
        assert!(s.doc().unwrap().path.is_none());
        assert!(s.doc().unwrap().display_title().ends_with(".xlsx"));
        assert_eq!(v(&s, "A1"), Value::Number(42.0));
        assert_eq!(v(&s, "B1"), Value::from("Original data"));
        assert!(s.doc().unwrap().wb.active().unwrap().cell(CellRef::parse("A1").unwrap()).unwrap().formula.is_none());
        s.execute("cell.set", json!({"cell": "B1", "input": "Changed"})).unwrap();
        assert_eq!(v(&s, "A1"), Value::Number(42.0)); // Imported formula caches are constants.
        let saved = s.execute("file.saveBytes", json!({"format": "xlsx"})).unwrap();
        s.execute("file.open", json!({"name": "imported.xlsx", "base64": saved["base64"]})).unwrap();
        assert_eq!(v(&s, "A1"), Value::Number(42.0));
    }
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn xlsb_import_never_reuses_source_as_save_target() {
    let nonce = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap().as_nanos();
    let dir = std::env::temp_dir().join(format!("gridcraft-xlsb-{}-{nonce}", std::process::id()));
    std::fs::create_dir(&dir).unwrap();
    let bytes = xlsb_fixture();
    for name in ["source.xlsb", "disguised.xlsx"] {
        let path = dir.join(name);
        std::fs::write(&path, &bytes).unwrap();
        let mut s = s();
        s.execute("file.open", json!({"path": path})).unwrap();
        s.take_ui_requests();
        s.execute("cell.set", json!({"cell": "A1", "input": "43"})).unwrap();
        s.execute("file.save", json!({})).unwrap();
        assert!(s.take_ui_requests().iter().any(|r| matches!(r, crate::UiRequest::Dialog(name, _) if name == "saveAs")));
        assert_eq!(std::fs::read(&path).unwrap(), bytes);
        assert!(s.execute("file.saveAs", json!({"path": dir.join("source.xlsb")})).is_err());
        assert!(s.execute("file.saveBytes", json!({"format": "xlsb"})).is_err());
        let output = dir.join("converted.xlsx");
        s.execute("file.saveAs", json!({"path": output})).unwrap();
        assert!(!s.doc().unwrap().is_dirty());
        assert_eq!(std::fs::read(&path).unwrap(), bytes);
        s.execute("file.open", json!({"name": "converted.xlsx", "base64": crate::io::base64_encode(&std::fs::read(output).unwrap())})).unwrap();
        assert_eq!(v(&s, "A1"), Value::Number(43.0));
    }
    std::fs::remove_dir_all(dir).unwrap();
}

#[test]
fn malformed_xlsb_keeps_the_current_workbook() {
    let mut s = s();
    s.execute("cell.set", json!({"cell": "A1", "input": "Keep me"})).unwrap();
    assert!(s.execute("file.open", json!({"name": "bad.xlsb", "base64": crate::io::base64_encode(b"not a workbook")})).is_err());
    assert_eq!(s.documents().len(), 1);
    assert_eq!(v(&s, "A1"), Value::from("Keep me"));
    assert!(s.doc().unwrap().is_dirty());
}
