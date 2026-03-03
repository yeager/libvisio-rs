//! Tests for SVG rendering functionality.

use libvisio_rs::model::*;
use libvisio_rs::svg::render;
use std::collections::HashMap;

// =============================================================================
// Color resolution tests
// =============================================================================

#[test]
fn test_resolve_color_empty() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("", &tc), "");
}

#[test]
fn test_resolve_color_hex() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("#FF0000", &tc), "#FF0000");
}

#[test]
fn test_resolve_color_hex_with_spaces() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("  #00FF00  ", &tc), "#00FF00");
}

#[test]
fn test_resolve_color_rgb_func() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("RGB(255,0,0)", &tc), "#FF0000");
}

#[test]
fn test_resolve_color_rgb_lowercase() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("rgb(0,128,255)", &tc), "#0080FF");
}

#[test]
fn test_resolve_color_rgb_with_spaces() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("RGB( 10 , 20 , 30 )", &tc), "#0A141E");
}

#[test]
fn test_resolve_color_index_0() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("0", &tc), "#000000");
}

#[test]
fn test_resolve_color_index_1() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("1", &tc), "#FFFFFF");
}

#[test]
fn test_resolve_color_index_2() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("2", &tc), "#FF0000");
}

#[test]
fn test_resolve_color_index_4() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("4", &tc), "#0000FF");
}

#[test]
fn test_resolve_color_index_14() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("14", &tc), "#C0C0C0");
}

#[test]
fn test_resolve_color_index_24() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("24", &tc), "#E6E6E6");
}

#[test]
fn test_resolve_color_inh() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("Inh", &tc), "");
}

#[test]
fn test_resolve_color_formula() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("=Sheet.1!Color", &tc), "");
}

#[test]
fn test_resolve_color_themeval_string() {
    let mut tc = HashMap::new();
    tc.insert("accent1".to_string(), "#4472C4".to_string());
    assert_eq!(
        render::resolve_color("THEMEVAL(\"accent1\",0)", &tc),
        "#4472C4"
    );
}

#[test]
fn test_resolve_color_themeval_numeric() {
    let mut tc = HashMap::new();
    tc.insert("4".to_string(), "#4472C4".to_string());
    assert_eq!(render::resolve_color("THEMEVAL(4)", &tc), "#4472C4");
}

#[test]
fn test_resolve_color_themeguard_themeval() {
    let mut tc = HashMap::new();
    tc.insert("accent2".to_string(), "#ED7D31".to_string());
    assert_eq!(
        render::resolve_color("THEMEGUARD(THEMEVAL(\"accent2\",0))", &tc),
        "#ED7D31"
    );
}

#[test]
fn test_resolve_color_themeval_missing_key() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("THEMEVAL(\"nonexistent\")", &tc), "");
}

#[test]
fn test_resolve_color_hsl() {
    let tc = HashMap::new();
    // HSL with Visio's 0-255 range
    let result = render::resolve_color("HSL(0,0,128)", &tc);
    assert!(result.starts_with('#'));
    assert_eq!(result.len(), 7);
}

#[test]
fn test_resolve_color_unknown_text() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_color("unknown_color", &tc), "");
}

// =============================================================================
// QuickStyle color resolution tests
// =============================================================================

#[test]
fn test_quickstyle_color_dk1() {
    let mut tc = HashMap::new();
    tc.insert("dk1".to_string(), "#000000".to_string());
    assert_eq!(render::resolve_quickstyle_color(0, &tc), "#000000");
}

#[test]
fn test_quickstyle_color_accent1() {
    let mut tc = HashMap::new();
    tc.insert("accent1".to_string(), "#4472C4".to_string());
    assert_eq!(render::resolve_quickstyle_color(2, &tc), "#4472C4");
    assert_eq!(render::resolve_quickstyle_color(102, &tc), "#4472C4");
}

#[test]
fn test_quickstyle_color_accent6() {
    let mut tc = HashMap::new();
    tc.insert("accent6".to_string(), "#70AD47".to_string());
    assert_eq!(render::resolve_quickstyle_color(7, &tc), "#70AD47");
}

#[test]
fn test_quickstyle_color_missing() {
    let tc = HashMap::new();
    assert_eq!(render::resolve_quickstyle_color(99, &tc), "");
}

// =============================================================================
// Color utility tests
// =============================================================================

#[test]
fn test_lighten_color() {
    let result = render::lighten_color("#000000", 0.5);
    assert_eq!(result, "#7F7F7F");
}

#[test]
fn test_lighten_color_white() {
    let result = render::lighten_color("#FFFFFF", 0.5);
    assert_eq!(result, "#FFFFFF");
}

#[test]
fn test_lighten_color_zero_factor() {
    let result = render::lighten_color("#808080", 0.0);
    assert_eq!(result, "#808080");
}

#[test]
fn test_lighten_color_full_factor() {
    let result = render::lighten_color("#000000", 1.0);
    assert_eq!(result, "#FFFFFF");
}

#[test]
fn test_lighten_color_invalid() {
    let result = render::lighten_color("bad", 0.5);
    assert_eq!(result, "#E8E8E8");
}

#[test]
fn test_is_dark_color_black() {
    assert!(render::is_dark_color("#000000"));
}

#[test]
fn test_is_dark_color_white() {
    assert!(!render::is_dark_color("#FFFFFF"));
}

#[test]
fn test_is_dark_color_dark_blue() {
    assert!(render::is_dark_color("#000080"));
}

#[test]
fn test_is_dark_color_light_gray() {
    assert!(!render::is_dark_color("#C0C0C0"));
}

#[test]
fn test_is_dark_color_empty() {
    assert!(!render::is_dark_color(""));
}

#[test]
fn test_is_dark_color_none() {
    assert!(!render::is_dark_color("none"));
}

// =============================================================================
// Shape model tests
// =============================================================================

#[test]
fn test_shape_default() {
    let shape = Shape::default();
    assert_eq!(shape.shape_type, "Shape");
    assert!(shape.text.is_empty());
    assert!(shape.geometry.is_empty());
    assert!(shape.sub_shapes.is_empty());
}

#[test]
fn test_shape_cell_val() {
    let mut shape = Shape::default();
    shape.cells.insert("Width".to_string(), CellValue::val("2.5"));
    assert_eq!(shape.cell_val("Width"), "2.5");
    assert_eq!(shape.cell_val("Missing"), "");
}

#[test]
fn test_shape_cell_f64() {
    let mut shape = Shape::default();
    shape.cells.insert("Width".to_string(), CellValue::val("2.5"));
    assert!((shape.cell_f64("Width") - 2.5).abs() < 1e-10);
    assert!((shape.cell_f64("Missing") - 0.0).abs() < 1e-10);
}

#[test]
fn test_shape_cell_f64_or() {
    let mut shape = Shape::default();
    shape.cells.insert("Width".to_string(), CellValue::val("3.0"));
    assert!((shape.cell_f64_or("Width", 1.0) - 3.0).abs() < 1e-10);
    assert!((shape.cell_f64_or("Missing", 5.0) - 5.0).abs() < 1e-10);
}

#[test]
fn test_cell_value_new() {
    let cv = CellValue::new("42", "=Width*2");
    assert_eq!(cv.v, "42");
    assert_eq!(cv.f, "=Width*2");
}

#[test]
fn test_cell_value_val() {
    let cv = CellValue::val("hello");
    assert_eq!(cv.v, "hello");
    assert!(cv.f.is_empty());
}

#[test]
fn test_cell_value_as_f64() {
    assert!((CellValue::val("3.15").as_f64() - 3.15).abs() < 1e-10);
    assert!((CellValue::val("bad").as_f64() - 0.0).abs() < 1e-10);
}

#[test]
fn test_cell_value_as_f64_or() {
    assert!((CellValue::val("3.15").as_f64_or(0.0) - 3.15).abs() < 1e-10);
    assert!((CellValue::val("bad").as_f64_or(42.0) - 42.0).abs() < 1e-10);
}

#[test]
fn test_cell_value_is_empty() {
    assert!(CellValue::default().is_empty());
    assert!(!CellValue::val("x").is_empty());
    assert!(!CellValue::new("", "=formula").is_empty());
}

#[test]
fn test_geom_row_new() {
    let row = GeomRow::new("MoveTo");
    assert_eq!(row.row_type, "MoveTo");
    assert!(row.cells.is_empty());
}

#[test]
fn test_geom_row_cell_f64() {
    let mut row = GeomRow::new("LineTo");
    row.cells.insert("X".to_string(), CellValue::val("1.5"));
    assert!((row.cell_f64("X") - 1.5).abs() < 1e-10);
    assert!((row.cell_f64("Y") - 0.0).abs() < 1e-10);
}

#[test]
fn test_geom_section_default() {
    let geo = GeomSection::default();
    assert!(!geo.no_fill);
    assert!(!geo.no_line);
    assert!(!geo.no_show);
    assert!(geo.rows.is_empty());
}

#[test]
fn test_char_format_default() {
    let cf = CharFormat::default();
    assert_eq!(cf.size, "0.1111");
    assert_eq!(cf.style, "0");
    assert!(cf.color.is_empty());
}

#[test]
fn test_page_default() {
    let page = Page::default();
    assert!((page.width - 8.5).abs() < 1e-10);
    assert!((page.height - 11.0).abs() < 1e-10);
    assert!(page.shapes.is_empty());
    assert!(!page.background);
}

#[test]
fn test_document_default() {
    let doc = Document::default();
    assert!(doc.pages.is_empty());
    assert!(doc.masters.is_empty());
    assert!(doc.theme_colors.is_empty());
}

// =============================================================================
// Connect model tests
// =============================================================================

#[test]
fn test_connect_fields() {
    let conn = Connect {
        from_sheet: "1".to_string(),
        from_cell: "BeginX".to_string(),
        to_sheet: "2".to_string(),
        to_cell: "PinX".to_string(),
    };
    assert_eq!(conn.from_sheet, "1");
    assert_eq!(conn.to_sheet, "2");
}

// =============================================================================
// Master merge tests
// =============================================================================

#[test]
fn test_merge_shape_with_master_no_master() {
    let mut shape = Shape::default();
    let masters = HashMap::new();
    render::merge_shape_with_master(&mut shape, &masters, "");
    // Should not panic
    assert!(shape.geometry.is_empty());
}

#[test]
fn test_merge_shape_with_master_cells() {
    let mut shape = Shape::default();
    shape.master = "1".to_string();
    shape.cells.insert("Width".to_string(), CellValue::val("3.0"));

    let mut master_shape = Shape::default();
    master_shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    master_shape.cells.insert("Height".to_string(), CellValue::val("1.5"));
    master_shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FF0000"));

    let mut master_shapes = HashMap::new();
    master_shapes.insert("default".to_string(), master_shape);
    let mut masters = HashMap::new();
    masters.insert("1".to_string(), master_shapes);

    render::merge_shape_with_master(&mut shape, &masters, "");

    // Shape's own Width should override master's
    assert_eq!(shape.cell_val("Width"), "3.0");
    // Master's Height should be merged in
    assert_eq!(shape.cell_val("Height"), "1.5");
    // Master's FillForegnd should be merged in
    assert_eq!(shape.cell_val("FillForegnd"), "#FF0000");
}

#[test]
fn test_merge_shape_with_master_geometry() {
    let mut shape = Shape::default();
    shape.master = "1".to_string();

    let mut master_shape = Shape::default();
    let geo = GeomSection {
        rows: vec![GeomRow::new("MoveTo")],
        ..GeomSection::default()
    };
    master_shape.geometry.push(geo);
    master_shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    master_shape.cells.insert("Height".to_string(), CellValue::val("1.0"));

    let mut master_shapes = HashMap::new();
    master_shapes.insert("default".to_string(), master_shape);
    let mut masters = HashMap::new();
    masters.insert("1".to_string(), master_shapes);

    render::merge_shape_with_master(&mut shape, &masters, "");

    assert_eq!(shape.geometry.len(), 1);
    assert!((shape.master_w - 2.0).abs() < 1e-10);
    assert!((shape.master_h - 1.0).abs() < 1e-10);
}

#[test]
fn test_merge_shape_with_master_text() {
    let mut shape = Shape::default();
    shape.master = "1".to_string();

    let mut master_shape = Shape::default();
    master_shape.text = "Master Text".to_string();

    let mut master_shapes = HashMap::new();
    master_shapes.insert("default".to_string(), master_shape);
    let mut masters = HashMap::new();
    masters.insert("1".to_string(), master_shapes);

    render::merge_shape_with_master(&mut shape, &masters, "");
    assert_eq!(shape.text, "Master Text");
}

#[test]
fn test_merge_shape_skips_placeholder_text() {
    let mut shape = Shape::default();
    shape.master = "1".to_string();

    let mut master_shape = Shape::default();
    master_shape.text = "Label".to_string();

    let mut master_shapes = HashMap::new();
    master_shapes.insert("default".to_string(), master_shape);
    let mut masters = HashMap::new();
    masters.insert("1".to_string(), master_shapes);

    render::merge_shape_with_master(&mut shape, &masters, "");
    assert!(shape.text.is_empty()); // "Label" should be skipped
}

#[test]
fn test_merge_shape_preserves_own_text() {
    let mut shape = Shape::default();
    shape.master = "1".to_string();
    shape.text = "My Text".to_string();

    let mut master_shape = Shape::default();
    master_shape.text = "Master Text".to_string();

    let mut master_shapes = HashMap::new();
    master_shapes.insert("default".to_string(), master_shape);
    let mut masters = HashMap::new();
    masters.insert("1".to_string(), master_shapes);

    render::merge_shape_with_master(&mut shape, &masters, "");
    assert_eq!(shape.text, "My Text"); // Should keep own text
}

#[test]
fn test_merge_shape_group_skips_text() {
    let mut shape = Shape::default();
    shape.master = "1".to_string();
    shape.shape_type = "Group".to_string();

    let mut master_shape = Shape::default();
    master_shape.text = "Master Text".to_string();

    let mut master_shapes = HashMap::new();
    master_shapes.insert("default".to_string(), master_shape);
    let mut masters = HashMap::new();
    masters.insert("1".to_string(), master_shapes);

    render::merge_shape_with_master(&mut shape, &masters, "");
    assert!(shape.text.is_empty()); // Groups don't inherit text
}

#[test]
fn test_merge_shape_parent_master_id() {
    let mut shape = Shape::default();
    // No master on shape, but parent master ID provided
    shape.cells.insert("Width".to_string(), CellValue::val("3.0"));

    let mut master_shape = Shape::default();
    master_shape.cells.insert("Height".to_string(), CellValue::val("2.0"));

    let mut master_shapes = HashMap::new();
    master_shapes.insert("default".to_string(), master_shape);
    let mut masters = HashMap::new();
    masters.insert("42".to_string(), master_shapes);

    render::merge_shape_with_master(&mut shape, &masters, "42");
    assert_eq!(shape.cell_val("Height"), "2.0");
}

// =============================================================================
// SVG rendering tests
// =============================================================================

fn empty_svg(shapes: &[Shape]) -> String {
    render::shapes_to_svg(
        shapes,
        8.5,
        11.0,
        &HashMap::new(),
        &[],
        &HashMap::new(),
        &HashMap::new(),
        None,
        &HashMap::new(),
        &HashMap::new(),
    )
}

#[test]
fn test_empty_page_svg() {
    let svg = empty_svg(&[]);
    assert!(svg.contains("<?xml"));
    assert!(svg.contains("<svg"));
    assert!(svg.contains("</svg>"));
    assert!(svg.contains("fill=\"white\""));
}

#[test]
fn test_simple_shape_svg() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.25"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.5"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#4472C4"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<rect"));
    assert!(svg.contains("#4472C4"));
}

#[test]
fn test_shape_with_geometry_svg() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FF0000"));

    let mut row1 = GeomRow::new("MoveTo");
    row1.cells.insert("X".to_string(), CellValue::val("0"));
    row1.cells.insert("Y".to_string(), CellValue::val("0"));
    let mut row2 = GeomRow::new("LineTo");
    row2.cells.insert("X".to_string(), CellValue::val("2.0"));
    row2.cells.insert("Y".to_string(), CellValue::val("0"));
    let mut row3 = GeomRow::new("LineTo");
    row3.cells.insert("X".to_string(), CellValue::val("2.0"));
    row3.cells.insert("Y".to_string(), CellValue::val("1.0"));
    let mut row4 = GeomRow::new("LineTo");
    row4.cells.insert("X".to_string(), CellValue::val("0"));
    row4.cells.insert("Y".to_string(), CellValue::val("1.0"));

    let geo = GeomSection {
        rows: vec![row1, row2, row3, row4],
        ..GeomSection::default()
    };
    shape.geometry.push(geo);

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<path"));
    assert!(svg.contains("M "));
    assert!(svg.contains("L "));
    assert!(svg.contains("#FF0000"));
}

#[test]
fn test_1d_connector_svg() {
    let mut shape = Shape::default();
    shape.id = "10".to_string();
    shape.cells.insert("BeginX".to_string(), CellValue::val("1.0"));
    shape.cells.insert("BeginY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("EndX".to_string(), CellValue::val("5.0"));
    shape.cells.insert("EndY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("EndArrow".to_string(), CellValue::val("4"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<line") || svg.contains("<path"));
    assert!(svg.contains("marker-end"));
    assert!(svg.contains("<marker"));
}

#[test]
fn test_1d_connector_orthogonal() {
    let mut shape = Shape::default();
    shape.id = "11".to_string();
    shape.cells.insert("BeginX".to_string(), CellValue::val("1.0"));
    shape.cells.insert("BeginY".to_string(), CellValue::val("8.0"));
    shape.cells.insert("EndX".to_string(), CellValue::val("5.0"));
    shape.cells.insert("EndY".to_string(), CellValue::val("3.0"));

    let svg = empty_svg(&[shape]);
    // Should generate orthogonal path (both axes differ significantly)
    assert!(svg.contains("<path"));
    assert!(svg.contains("stroke-linejoin=\"round\""));
}

#[test]
fn test_invisible_shape_skipped() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("Visible".to_string(), CellValue::val("0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.text = "Should not appear".to_string();

    let svg = empty_svg(&[shape]);
    assert!(!svg.contains("Should not appear"));
}

#[test]
fn test_text_rendering() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.text = "Hello World".to_string();

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<text"));
    assert!(svg.contains("Hello World"));
}

#[test]
fn test_text_xml_escaping() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.text = "A < B & C > D".to_string();

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("A &lt; B &amp; C &gt; D"));
}

#[test]
fn test_multiline_text() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.text = "Line 1\nLine 2\nLine 3".to_string();

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("Line 1"));
    assert!(svg.contains("Line 2"));
    assert!(svg.contains("Line 3"));
    // Should have multiple text elements
    let count = svg.matches("<text").count();
    assert!(count >= 3);
}

#[test]
fn test_hatching_pattern_fill() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillPattern".to_string(), CellValue::val("3"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#0000FF"));
    shape.cells.insert("FillBkgnd".to_string(), CellValue::val("#FFFFFF"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<pattern"));
    assert!(svg.contains("url(#fpat_1_3)"));
}

#[test]
fn test_gradient_fill() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillPattern".to_string(), CellValue::val("25"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FF0000"));
    shape.cells.insert("FillBkgnd".to_string(), CellValue::val("#0000FF"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<linearGradient") || svg.contains("<radialGradient"));
    assert!(svg.contains("url(#grad_1_25)"));
}

#[test]
fn test_radial_gradient_fill() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillPattern".to_string(), CellValue::val("29"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FF0000"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<radialGradient"));
}

#[test]
fn test_dashed_line() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("LinePattern".to_string(), CellValue::val("2"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FFFFFF"));

    let mut row1 = GeomRow::new("MoveTo");
    row1.cells.insert("X".to_string(), CellValue::val("0"));
    row1.cells.insert("Y".to_string(), CellValue::val("0"));
    let mut row2 = GeomRow::new("LineTo");
    row2.cells.insert("X".to_string(), CellValue::val("2.0"));
    row2.cells.insert("Y".to_string(), CellValue::val("1.0"));
    let geo = GeomSection {
        rows: vec![row1, row2],
        ..GeomSection::default()
    };
    shape.geometry.push(geo);

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("stroke-dasharray"));
}

#[test]
fn test_no_line_pattern() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("LinePattern".to_string(), CellValue::val("0"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FF0000"));

    let mut row1 = GeomRow::new("MoveTo");
    row1.cells.insert("X".to_string(), CellValue::val("0"));
    row1.cells.insert("Y".to_string(), CellValue::val("0"));
    let mut row2 = GeomRow::new("LineTo");
    row2.cells.insert("X".to_string(), CellValue::val("2.0"));
    row2.cells.insert("Y".to_string(), CellValue::val("1.0"));
    let geo = GeomSection {
        rows: vec![row1, row2],
        ..GeomSection::default()
    };
    shape.geometry.push(geo);

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("stroke=\"none\""));
}

#[test]
fn test_no_fill_geometry() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillPattern".to_string(), CellValue::val("0"));

    let mut row1 = GeomRow::new("MoveTo");
    row1.cells.insert("X".to_string(), CellValue::val("0"));
    row1.cells.insert("Y".to_string(), CellValue::val("0"));
    let mut row2 = GeomRow::new("LineTo");
    row2.cells.insert("X".to_string(), CellValue::val("2.0"));
    row2.cells.insert("Y".to_string(), CellValue::val("1.0"));
    let geo = GeomSection {
        rows: vec![row1, row2],
        ..GeomSection::default()
    };
    shape.geometry.push(geo);

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("fill=\"none\""));
}

#[test]
fn test_group_shape() {
    let mut group = Shape::default();
    group.id = "1".to_string();
    group.shape_type = "Group".to_string();
    group.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    group.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    group.cells.insert("Width".to_string(), CellValue::val("3.0"));
    group.cells.insert("Height".to_string(), CellValue::val("2.0"));

    let mut child = Shape::default();
    child.id = "2".to_string();
    child.cells.insert("PinX".to_string(), CellValue::val("1.5"));
    child.cells.insert("PinY".to_string(), CellValue::val("1.0"));
    child.cells.insert("Width".to_string(), CellValue::val("1.0"));
    child.cells.insert("Height".to_string(), CellValue::val("0.5"));
    child.cells.insert("FillForegnd".to_string(), CellValue::val("#00FF00"));
    group.sub_shapes.push(child);

    let svg = empty_svg(&[group]);
    assert!(svg.contains("<g"));
    assert!(svg.contains("</g>"));
}

#[test]
fn test_fill_opacity() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FF0000"));
    shape.cells.insert("FillForegndTrans".to_string(), CellValue::val("0.5"));

    let mut row1 = GeomRow::new("MoveTo");
    row1.cells.insert("X".to_string(), CellValue::val("0"));
    row1.cells.insert("Y".to_string(), CellValue::val("0"));
    let mut row2 = GeomRow::new("LineTo");
    row2.cells.insert("X".to_string(), CellValue::val("2.0"));
    row2.cells.insert("Y".to_string(), CellValue::val("1.0"));
    let geo = GeomSection {
        rows: vec![row1, row2],
        ..GeomSection::default()
    };
    shape.geometry.push(geo);

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("fill-opacity=\"0.50\""));
}

#[test]
fn test_background_shapes() {
    let mut fg_shape = Shape::default();
    fg_shape.id = "1".to_string();
    fg_shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    fg_shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    fg_shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    fg_shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    fg_shape.text = "Foreground".to_string();

    let mut bg_shape = Shape::default();
    bg_shape.id = "2".to_string();
    bg_shape.cells.insert("PinX".to_string(), CellValue::val("4.25"));
    bg_shape.cells.insert("PinY".to_string(), CellValue::val("5.5"));
    bg_shape.cells.insert("Width".to_string(), CellValue::val("8.5"));
    bg_shape.cells.insert("Height".to_string(), CellValue::val("11.0"));
    bg_shape.cells.insert("FillForegnd".to_string(), CellValue::val("#EEEEEE"));

    let bg_shapes = vec![bg_shape];
    let svg = render::shapes_to_svg(
        &[fg_shape],
        8.5,
        11.0,
        &HashMap::new(),
        &[],
        &HashMap::new(),
        &HashMap::new(),
        Some(&bg_shapes),
        &HashMap::new(),
        &HashMap::new(),
    );
    assert!(svg.contains("Background page"));
    assert!(svg.contains("Foreground"));
}

#[test]
fn test_layer_visibility() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("LayerMember".to_string(), CellValue::val("0"));
    shape.text = "Hidden by layer".to_string();

    let mut layers = HashMap::new();
    layers.insert(
        "0".to_string(),
        LayerDef {
            name: "Layer 0".to_string(),
            visible: false,
        },
    );

    let svg = render::shapes_to_svg(
        &[shape],
        8.5,
        11.0,
        &HashMap::new(),
        &[],
        &HashMap::new(),
        &HashMap::new(),
        None,
        &HashMap::new(),
        &layers,
    );
    assert!(!svg.contains("Hidden by layer"));
}

// =============================================================================
// Geometry type tests
// =============================================================================

#[test]
fn test_ellipse_geometry() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("2.0"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FF0000"));

    let mut row = GeomRow::new("Ellipse");
    row.cells.insert("X".to_string(), CellValue::val("1.0"));
    row.cells.insert("Y".to_string(), CellValue::val("1.0"));
    row.cells.insert("A".to_string(), CellValue::val("2.0"));
    row.cells.insert("B".to_string(), CellValue::val("1.0"));
    row.cells.insert("C".to_string(), CellValue::val("1.0"));
    row.cells.insert("D".to_string(), CellValue::val("2.0"));
    let geo = GeomSection {
        rows: vec![row],
        ..GeomSection::default()
    };
    shape.geometry.push(geo);

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<path"));
    assert!(svg.contains("A ")); // Arc command for ellipse
}

#[test]
fn test_arc_geometry() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));

    let mut row1 = GeomRow::new("MoveTo");
    row1.cells.insert("X".to_string(), CellValue::val("0"));
    row1.cells.insert("Y".to_string(), CellValue::val("0.5"));
    let mut row2 = GeomRow::new("ArcTo");
    row2.cells.insert("X".to_string(), CellValue::val("2.0"));
    row2.cells.insert("Y".to_string(), CellValue::val("0.5"));
    row2.cells.insert("A".to_string(), CellValue::val("0.3"));
    let geo = GeomSection {
        rows: vec![row1, row2],
        ..GeomSection::default()
    };
    shape.geometry.push(geo);

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<path"));
    assert!(svg.contains("A ")); // Arc command
}

#[test]
fn test_no_show_geometry_skipped() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#FF0000"));

    let mut row1 = GeomRow::new("MoveTo");
    row1.cells.insert("X".to_string(), CellValue::val("0"));
    row1.cells.insert("Y".to_string(), CellValue::val("0"));
    let geo = GeomSection {
        no_show: true,
        rows: vec![row1],
        ..GeomSection::default()
    };
    shape.geometry.push(geo);

    let svg = empty_svg(&[shape]);
    // Should not have a path element (geometry is hidden)
    assert!(!svg.contains("<path"));
}

// =============================================================================
// Connector label tests
// =============================================================================

#[test]
fn test_connector_label_background() {
    let mut shape = Shape::default();
    shape.id = "20".to_string();
    shape.cells.insert("BeginX".to_string(), CellValue::val("1.0"));
    shape.cells.insert("BeginY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("EndX".to_string(), CellValue::val("5.0"));
    shape.cells.insert("EndY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("PinX".to_string(), CellValue::val("3.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.text = "Connection".to_string();

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("Connection"));
    // Should have white background rect for readability
    assert!(svg.contains("fill-opacity=\"0.85\""));
}

// =============================================================================
// Hatching pattern type tests
// =============================================================================

#[test]
fn test_hatching_horizontal_lines() {
    let mut shape = Shape::default();
    shape.id = "h2".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillPattern".to_string(), CellValue::val("2"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#000000"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<pattern id=\"fpat_h2_2\""));
    assert!(svg.contains("url(#fpat_h2_2)"));
}

#[test]
fn test_hatching_crosshatch() {
    let mut shape = Shape::default();
    shape.id = "h6".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillPattern".to_string(), CellValue::val("6"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#000000"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<pattern id=\"fpat_h6_6\""));
}

#[test]
fn test_hatching_diagonal_crosshatch() {
    let mut shape = Shape::default();
    shape.id = "h10".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillPattern".to_string(), CellValue::val("10"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#000000"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<pattern id=\"fpat_h10_10\""));
}

#[test]
fn test_hatching_dense_dots() {
    let mut shape = Shape::default();
    shape.id = "h15".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert("FillPattern".to_string(), CellValue::val("15"));
    shape.cells.insert("FillForegnd".to_string(), CellValue::val("#000000"));

    let svg = empty_svg(&[shape]);
    assert!(svg.contains("<pattern id=\"fpat_h15_15\""));
    assert!(svg.contains("<circle"));
}

// =============================================================================
// Theme color tests
// =============================================================================

#[test]
fn test_theme_color_in_fill() {
    let mut shape = Shape::default();
    shape.id = "1".to_string();
    shape.cells.insert("PinX".to_string(), CellValue::val("4.0"));
    shape.cells.insert("PinY".to_string(), CellValue::val("5.0"));
    shape.cells.insert("Width".to_string(), CellValue::val("2.0"));
    shape.cells.insert("Height".to_string(), CellValue::val("1.0"));
    shape.cells.insert(
        "FillForegnd".to_string(),
        CellValue::new("#4472C4", "THEMEVAL(\"accent1\",0)"),
    );

    let mut theme_colors = HashMap::new();
    theme_colors.insert("accent1".to_string(), "#5B9BD5".to_string());

    let svg = render::shapes_to_svg(
        &[shape],
        8.5,
        11.0,
        &HashMap::new(),
        &[],
        &HashMap::new(),
        &HashMap::new(),
        None,
        &theme_colors,
        &HashMap::new(),
    );
    // Theme color should be resolved
    assert!(svg.contains("#5B9BD5") || svg.contains("#4472C4"));
}

// =============================================================================
// Page info / gradient stop model tests
// =============================================================================

#[test]
fn test_page_info() {
    let pi = PageInfo {
        name: "Page 1".to_string(),
        index: 0,
        width: 8.5,
        height: 11.0,
    };
    assert_eq!(pi.name, "Page 1");
    assert_eq!(pi.index, 0);
}

#[test]
fn test_gradient_stop() {
    let stop = GradientStop {
        position: 50.0,
        color: "#FF0000".to_string(),
    };
    assert!((stop.position - 50.0).abs() < 1e-10);
}

#[test]
fn test_gradient_def_default() {
    let gd = GradientDef::default();
    assert_eq!(gd.start, "#FFFFFF");
    assert_eq!(gd.end, "#000000");
    assert!(!gd.radial);
}

#[test]
fn test_fill_pattern_def() {
    let fpd = FillPatternDef {
        id: "test".to_string(),
        fg: "#000000".to_string(),
        bg: "#FFFFFF".to_string(),
        pattern_type: 3,
    };
    assert_eq!(fpd.pattern_type, 3);
}

#[test]
fn test_hyperlink() {
    let hl = Hyperlink {
        description: "Test".to_string(),
        address: "https://example.com".to_string(),
        sub_address: String::new(),
        frame: String::new(),
    };
    assert_eq!(hl.address, "https://example.com");
}

#[test]
fn test_stylesheet_default() {
    let ss = StyleSheet::default();
    assert!(ss.cells.is_empty());
}

#[test]
fn test_layer_def() {
    let ld = LayerDef {
        name: "Connector".to_string(),
        visible: true,
    };
    assert!(ld.visible);
}

#[test]
fn test_foreign_data_info_default() {
    let fdi = ForeignDataInfo::default();
    assert!(fdi.foreign_type.is_empty());
    assert!(fdi.data.is_none());
}

#[test]
fn test_para_format_default() {
    let pf = ParaFormat::default();
    assert!(pf.horiz_align.is_empty());
}

#[test]
fn test_text_part() {
    let tp = TextPart {
        text: "Hello".to_string(),
        cp: "0".to_string(),
        pp: "0".to_string(),
    };
    assert_eq!(tp.text, "Hello");
}

#[test]
fn test_xform_default() {
    let xf = XForm::default();
    assert!((xf.pin_x - 0.0).abs() < 1e-10);
    assert!(!xf.flip_x);
}

#[test]
fn test_xform1d_default() {
    let xf = XForm1D::default();
    assert!((xf.begin_x - 0.0).abs() < 1e-10);
    assert!((xf.end_x - 0.0).abs() < 1e-10);
}

#[test]
fn test_text_xform_default() {
    let txf = TextXForm::default();
    assert!((txf.txt_pin_x - 0.0).abs() < 1e-10);
}
