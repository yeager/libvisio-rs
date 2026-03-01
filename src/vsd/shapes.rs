//! Shape/group hierarchy conversion for .vsd binary shapes.
//!
//! Converts VSD binary parsed data into the common Shape model.

use crate::model::*;
use base64::Engine;
use std::collections::HashMap;

/// Convert VSD binary shape data to the common Shape model dict format.
/// This is called by the VSD parser after reading the binary stream.
pub fn vsd_shape_to_model(
    xform: &VsdXForm,
    text: &str,
    geometry: &[VsdGeomSection],
    line_weight: f64,
    line_color: &str,
    line_pattern: i32,
    fill_fg: &str,
    fill_bg: &str,
    fill_pattern: i32,
    shape_id: u32,
    shape_type: &str,
    text_xform: Option<&VsdTextXForm>,
    xform_1d: Option<&VsdXForm1D>,
    char_formats: &[VsdCharFormat],
    para_formats: &[VsdParaFormat],
    shadow_color: &str,
    shadow_pattern: i32,
    shadow_offset_x: f64,
    shadow_offset_y: f64,
    foreign_data: Option<&VsdForeignData>,
    layer_member: &str,
) -> Shape {
    let mut shape = Shape::default();
    shape.id = shape_id.to_string();
    shape.shape_type = shape_type.to_string();

    // XForm cells
    shape
        .cells
        .insert("PinX".to_string(), CellValue::val(xform.pin_x.to_string()));
    shape
        .cells
        .insert("PinY".to_string(), CellValue::val(xform.pin_y.to_string()));
    shape
        .cells
        .insert("Width".to_string(), CellValue::val(xform.width.to_string()));
    shape.cells.insert(
        "Height".to_string(),
        CellValue::val(xform.height.to_string()),
    );
    shape.cells.insert(
        "LocPinX".to_string(),
        CellValue::val(xform.loc_pin_x.to_string()),
    );
    shape.cells.insert(
        "LocPinY".to_string(),
        CellValue::val(xform.loc_pin_y.to_string()),
    );
    shape
        .cells
        .insert("Angle".to_string(), CellValue::val(xform.angle.to_string()));
    if xform.flip_x {
        shape.cells.insert("FlipX".to_string(), CellValue::val("1"));
    }
    if xform.flip_y {
        shape.cells.insert("FlipY".to_string(), CellValue::val("1"));
    }

    // Line/fill cells
    shape.cells.insert(
        "LineWeight".to_string(),
        CellValue::val(line_weight.to_string()),
    );
    shape
        .cells
        .insert("LineColor".to_string(), CellValue::val(line_color));
    shape.cells.insert(
        "LinePattern".to_string(),
        CellValue::val(line_pattern.to_string()),
    );
    if !fill_fg.is_empty() {
        shape
            .cells
            .insert("FillForegnd".to_string(), CellValue::val(fill_fg));
    }
    if !fill_bg.is_empty() {
        shape
            .cells
            .insert("FillBkgnd".to_string(), CellValue::val(fill_bg));
    }
    shape.cells.insert(
        "FillPattern".to_string(),
        CellValue::val(fill_pattern.to_string()),
    );

    // Text
    shape.text = text.to_string();
    if !text.is_empty() {
        shape.text_parts.push(TextPart {
            text: text.to_string(),
            cp: "0".to_string(),
            pp: "0".to_string(),
        });
    }

    // TextXForm
    if let Some(txf) = text_xform {
        shape.cells.insert(
            "TxtPinX".to_string(),
            CellValue::val(txf.txt_pin_x.to_string()),
        );
        shape.cells.insert(
            "TxtPinY".to_string(),
            CellValue::val(txf.txt_pin_y.to_string()),
        );
        shape.cells.insert(
            "TxtWidth".to_string(),
            CellValue::val(txf.txt_width.to_string()),
        );
        shape.cells.insert(
            "TxtHeight".to_string(),
            CellValue::val(txf.txt_height.to_string()),
        );
    }

    // XForm1D
    if let Some(xf1d) = xform_1d {
        shape.cells.insert(
            "BeginX".to_string(),
            CellValue::val(xf1d.begin_x.to_string()),
        );
        shape.cells.insert(
            "BeginY".to_string(),
            CellValue::val(xf1d.begin_y.to_string()),
        );
        shape
            .cells
            .insert("EndX".to_string(), CellValue::val(xf1d.end_x.to_string()));
        shape
            .cells
            .insert("EndY".to_string(), CellValue::val(xf1d.end_y.to_string()));
    }

    // Shadow
    if !shadow_color.is_empty() {
        shape
            .cells
            .insert("ShdwForegnd".to_string(), CellValue::val(shadow_color));
    }
    if shadow_pattern != 0 {
        shape.cells.insert(
            "ShdwPattern".to_string(),
            CellValue::val(shadow_pattern.to_string()),
        );
    }
    if shadow_offset_x != 0.0 {
        shape.cells.insert(
            "ShdwOffsetX".to_string(),
            CellValue::val(shadow_offset_x.to_string()),
        );
    }
    if shadow_offset_y != 0.0 {
        shape.cells.insert(
            "ShdwOffsetY".to_string(),
            CellValue::val(shadow_offset_y.to_string()),
        );
    }

    // Layer
    if !layer_member.is_empty() {
        shape
            .cells
            .insert("LayerMember".to_string(), CellValue::val(layer_member));
    }

    // Character formats
    for (i, cf) in char_formats.iter().enumerate() {
        shape.char_formats.insert(
            i.to_string(),
            CharFormat {
                size: if cf.font_size > 0.0 {
                    (cf.font_size / 72.0).to_string()
                } else {
                    "0.1111".to_string()
                },
                color: format!("#{:02X}{:02X}{:02X}", cf.color_r, cf.color_g, cf.color_b),
                style: ((if cf.bold { 1 } else { 0 })
                    | (if cf.italic { 2 } else { 0 })
                    | (if cf.underline { 4 } else { 0 }))
                .to_string(),
                font: String::new(),
            },
        );
    }

    // Paragraph formats
    for (i, pf) in para_formats.iter().enumerate() {
        shape.para_formats.insert(
            i.to_string(),
            ParaFormat {
                horiz_align: pf.horiz_align.to_string(),
                indent_first: pf.indent_first.to_string(),
                indent_left: pf.indent_left.to_string(),
                indent_right: pf.indent_right.to_string(),
                bullet: pf.bullet.to_string(),
                sp_line: pf.spacing_line.to_string(),
                ..ParaFormat::default()
            },
        );
    }

    // Geometry
    for geom in geometry {
        let mut geo = GeomSection {
            no_fill: geom.no_fill,
            no_line: geom.no_line,
            no_show: geom.no_show,
            ix: "0".to_string(),
            rows: Vec::new(),
        };
        for row in &geom.rows {
            let mut cells = HashMap::new();
            cells.insert("X".to_string(), CellValue::val(row.x.to_string()));
            cells.insert("Y".to_string(), CellValue::val(row.y.to_string()));
            if matches!(
                row.row_type.as_str(),
                "ArcTo"
                    | "Ellipse"
                    | "EllipticalArcTo"
                    | "SplineStart"
                    | "SplineKnot"
                    | "InfiniteLine"
            ) {
                cells.insert("A".to_string(), CellValue::val(row.a.to_string()));
                cells.insert("B".to_string(), CellValue::val(row.b.to_string()));
            }
            if matches!(
                row.row_type.as_str(),
                "Ellipse" | "EllipticalArcTo" | "SplineStart"
            ) {
                cells.insert("C".to_string(), CellValue::val(row.c.to_string()));
                cells.insert("D".to_string(), CellValue::val(row.d.to_string()));
            }
            if row.row_type == "NURBSTo" {
                // Encode control points as semicolon-separated values
                if !row.points.is_empty() {
                    let pts_str: String = row
                        .points
                        .iter()
                        .map(|p| format!("{},{},{},{}", p.0, p.1, p.2, p.3))
                        .collect::<Vec<_>>()
                        .join(";");
                    cells.insert("E".to_string(), CellValue::new(pts_str, "NURBS(...)"));
                }
            }
            if row.row_type == "PolylineTo" {
                if !row.points.is_empty() {
                    let pts_str: String = row
                        .points
                        .iter()
                        .map(|p| format!("{},{}", p.0, p.1))
                        .collect::<Vec<_>>()
                        .join(";");
                    cells.insert("A".to_string(), CellValue::new(pts_str, "POLYLINE(...)"));
                }
            }
            geo.rows.push(GeomRow {
                row_type: row.row_type.clone(),
                ix: String::new(),
                cells,
            });
        }
        shape.geometry.push(geo);
    }

    // Foreign data (embedded images)
    if let Some(fd) = foreign_data {
        if !fd.data.is_empty() {
            let b64 = base64::engine::general_purpose::STANDARD.encode(&fd.data);
            shape.foreign_data = Some(ForeignDataInfo {
                foreign_type: fd.data_type.clone(),
                compression: fd.img_format.clone(),
                data: Some(b64),
                rel_id: None,
            });
        }
    }

    shape
}

// VSD-specific data types used during parsing

#[derive(Debug, Clone, Default)]
pub struct VsdXForm {
    pub pin_x: f64,
    pub pin_y: f64,
    pub width: f64,
    pub height: f64,
    pub loc_pin_x: f64,
    pub loc_pin_y: f64,
    pub angle: f64,
    pub flip_x: bool,
    pub flip_y: bool,
}

#[derive(Debug, Clone, Default)]
pub struct VsdTextXForm {
    pub txt_pin_x: f64,
    pub txt_pin_y: f64,
    pub txt_width: f64,
    pub txt_height: f64,
}

#[derive(Debug, Clone, Default)]
pub struct VsdXForm1D {
    pub begin_x: f64,
    pub begin_y: f64,
    pub end_x: f64,
    pub end_y: f64,
}

#[derive(Debug, Clone, Default)]
pub struct VsdGeomRow {
    pub row_type: String,
    pub x: f64,
    pub y: f64,
    pub a: f64,
    pub b: f64,
    pub c: f64,
    pub d: f64,
    pub knot_last: f64,
    pub degree: u16,
    pub x_type: u8,
    pub y_type: u8,
    pub points: Vec<(f64, f64, f64, f64)>, // (x, y, knot, weight) for NURBS; (x, y, 0, 0) for polyline
}

#[derive(Debug, Clone, Default)]
pub struct VsdGeomSection {
    pub no_fill: bool,
    pub no_line: bool,
    pub no_show: bool,
    pub rows: Vec<VsdGeomRow>,
}

#[derive(Debug, Clone, Default)]
pub struct VsdCharFormat {
    pub char_count: u32,
    pub font_id: u16,
    pub color_r: u8,
    pub color_g: u8,
    pub color_b: u8,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub font_size: f64,
}

#[derive(Debug, Clone, Default)]
pub struct VsdParaFormat {
    pub char_count: u32,
    pub indent_first: f64,
    pub indent_left: f64,
    pub indent_right: f64,
    pub spacing_line: f64,
    pub horiz_align: u8,
    pub bullet: u8,
}

#[derive(Debug, Clone, Default)]
pub struct VsdForeignData {
    pub data_type: String,
    pub img_format: String,
    pub data: Vec<u8>,
}
