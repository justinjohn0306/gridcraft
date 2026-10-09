use gridcraft_engine::Session;
use serde_json::json;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut session = Session::new();
    session.new_workbook();

    // Add data: months and sales figures with some curves
    let months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    let sales = [45.0, 52.0, 48.0, 65.0, 58.0, 72.0, 68.0, 75.0, 70.0, 82.0, 78.0, 85.0];
    let expenses = [30.0, 35.0, 32.0, 42.0, 38.0, 45.0, 43.0, 48.0, 46.0, 52.0, 50.0, 55.0];

    // Headers
    session.execute("cell.set", json!({"cell": "A1", "input": "Month"}))?;
    session.execute("cell.set", json!({"cell": "B1", "input": "Sales"}))?;
    session.execute("cell.set", json!({"cell": "C1", "input": "Expenses"}))?;

    // Data
    for (i, ((month, sale), expense)) in months.iter().zip(&sales).zip(&expenses).enumerate() {
        let row = i + 2;
        session.execute("cell.set", json!({"cell": format!("A{}", row), "input": month}))?;
        session.execute("cell.set", json!({"cell": format!("B{}", row), "input": sale.to_string()}))?;
        session.execute("cell.set", json!({"cell": format!("C{}", row), "input": expense.to_string()}))?;
    }

    // Create first chart (straight lines)
    session.execute(
        "insert.chart",
        json!({
            "type": "line",
            "subtype": "markers",
            "range": "A1:C13",
            "title": "Straight Lines (default)",
            "at": "E2"
        }),
    )?;

    // Create second chart (smooth lines)
    session.execute(
        "insert.chart",
        json!({
            "type": "line",
            "subtype": "markers",
            "range": "A1:C13",
            "title": "Smooth Curves",
            "at": "E18"
        }),
    )?;

    // Enable smoothing on the second chart
    session.execute(
        "chart.addElement",
        json!({
            "chart": 2,
            "element": "smooth",
            "on": true
        }),
    )?;

    // Save the file
    session.execute("file.saveAs", json!({"path": "smooth_demo.xlsx"}))?;

    println!("Created smooth_demo.xlsx with two charts:");
    println!("  - Chart 1: Straight lines (smooth=false)");
    println!("  - Chart 2: Smooth curves (smooth=true)");
    println!("\nTo view it:");
    println!("  cargo run --release -p gridcraft -- smooth_demo.xlsx");

    Ok(())
}
