//! Shape → SVG conversion.
//!
//! This module contains the complete SVG rendering pipeline, ported from
//! the Python libvisio-ng. It handles geometry paths, text, gradients,
//! shadows, arrows, fills, and all the visual features.

use crate::model::*;
use std::collections::{HashMap, HashSet};

/// Inches to SVG pixels.
const INCH_TO_PX: f64 = 72.0;

/// Visio color index table.
static VISIO_COLORS: &[(i32, &str)] = &[
    (0, "#000000"),
    (1, "#FFFFFF"),
    (2, "#FF0000"),
    (3, "#00FF00"),
    (4, "#0000FF"),
    (5, "#FFFF00"),
    (6, "#FF00FF"),
    (7, "#00FFFF"),
    (8, "#800000"),
    (9, "#008000"),
    (10, "#000080"),
    (11, "#808000"),
    (12, "#800080"),
    (13, "#008080"),
    (14, "#C0C0C0"),
    (15, "#808080"),
    (16, "#993366"),
    (17, "#333399"),
    (18, "#333333"),
    (19, "#003300"),
    (20, "#003366"),
    (21, "#993300"),
    (22, "#993366"),
    (23, "#333399"),
    (24, "#E6E6E6"),
];

/// Line pattern dash arrays.
static LINE_PATTERNS: &[(i32, &str)] = &[
    (0, "none"),
    (1, ""),
    (2, "4,3"),
    (3, "1,3"),
    (4, "4,3,1,3"),
    (5, "4,3,1,3,1,3"),
    (6, "8,3"),
    (7, "1,1"),
    (8, "8,3,1,3"),
    (9, "8,3,1,3,1,3"),
    (10, "12,6"),
    (16, "6,3,6,3"),
];

/// Arrow size scale factors.
static ARROW_SIZES: &[(i32, f64)] = &[
    (0, 0.6),
    (1, 0.8),
    (2, 1.0),
    (3, 1.2),
    (4, 1.6),
    (5, 2.0),
    (6, 2.5),
];

fn arrow_size(idx: i32) -> f64 {
    ARROW_SIZES
        .iter()
        .find(|&&(i, _)| i == idx)
        .map(|&(_, s)| s)
        .unwrap_or(1.0)
}

/// Resolve a Visio color value to SVG hex.
pub fn resolve_color(val: &str, theme_colors: &HashMap<String, String>) -> String {
    let val = val.trim();
    if val.is_empty() {
        return String::new();
    }

    // THEMEVAL
    if val.contains("THEMEVAL") || val.contains("THEMEGUARD") {
        if let Some(key) = extract_themeval_key(val) {
            if let Some(color) = theme_colors.get(&key) {
                return color.clone();
            }
        }
        return String::new();
    }

    if val == "Inh" || val.starts_with('=') || val.contains("THEME") {
        return String::new();
    }

    // #RRGGBB
    if val.starts_with('#') {
        return val.to_string();
    }

    // RGB(r,g,b)
    if let Some(color) = parse_rgb_func(val) {
        return color;
    }

    // HSL(h,s,l)
    if let Some(color) = parse_hsl_func(val) {
        return color;
    }

    // Numeric index
    if let Ok(idx) = val.parse::<f64>() {
        let i = idx as i32;
        if let Some(&(_, color)) = VISIO_COLORS.iter().find(|&&(ci, _)| ci == i) {
            return color.to_string();
        }
    }

    String::new()
}

fn extract_themeval_key(val: &str) -> Option<String> {
    // THEMEVAL("accent1",0) or THEMEVAL(0)
    let upper = val.to_uppercase();
    if let Some(start) = upper.find("THEMEVAL") {
        let rest = &val[start + 8..];
        if let Some(paren_start) = rest.find('(') {
            let inner = &rest[paren_start + 1..];
            // Try quoted key
            if let Some(q_start) = inner.find('"') {
                if let Some(q_end) = inner[q_start + 1..].find('"') {
                    return Some(inner[q_start + 1..q_start + 1 + q_end].to_lowercase());
                }
            }
            // Try numeric
            let num_str: String = inner.chars().take_while(|c| c.is_ascii_digit()).collect();
            if !num_str.is_empty() {
                return Some(num_str);
            }
        }
    }
    None
}

fn parse_rgb_func(val: &str) -> Option<String> {
    let upper = val.to_uppercase();
    if !upper.starts_with("RGB") {
        return None;
    }
    let inner = val.split('(').nth(1)?.split(')').next()?;
    let parts: Vec<&str> = inner.split(',').collect();
    if parts.len() != 3 {
        return None;
    }
    let r: u8 = parts[0].trim().parse().ok()?;
    let g: u8 = parts[1].trim().parse().ok()?;
    let b: u8 = parts[2].trim().parse().ok()?;
    Some(format!("#{:02X}{:02X}{:02X}", r, g, b))
}

fn parse_hsl_func(val: &str) -> Option<String> {
    let upper = val.to_uppercase();
    if !upper.starts_with("HSL") {
        return None;
    }
    let inner = val.split('(').nth(1)?.split(')').next()?;
    let parts: Vec<&str> = inner.split(',').collect();
    if parts.len() != 3 {
        return None;
    }
    let h: f64 = parts[0].trim().parse().ok()?;
    let s: f64 = parts[1].trim().parse().ok()?;
    let l: f64 = parts[2].trim().parse().ok()?;
    Some(hsl_to_rgb(h as i32, s as i32, l as i32))
}

fn hsl_to_rgb(h: i32, s: i32, l: i32) -> String {
    let hf = (h as f64 / 255.0) * 360.0;
    let sf = s as f64 / 255.0;
    let lf = l as f64 / 255.0;
    let (r, g, b) = if sf == 0.0 {
        (lf, lf, lf)
    } else {
        let q = if lf < 0.5 {
            lf * (1.0 + sf)
        } else {
            lf + sf - lf * sf
        };
        let p = 2.0 * lf - q;
        let hn = hf / 360.0;
        (
            hue2rgb(p, q, hn + 1.0 / 3.0),
            hue2rgb(p, q, hn),
            hue2rgb(p, q, hn - 1.0 / 3.0),
        )
    };
    format!(
        "#{:02X}{:02X}{:02X}",
        (r * 255.0) as u8,
        (g * 255.0) as u8,
        (b * 255.0) as u8
    )
}

fn hue2rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }
    if t < 1.0 / 6.0 {
        return p + (q - p) * 6.0 * t;
    }
    if t < 0.5 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
    }
    p
}

#[allow(dead_code)]
fn lighten_color(hex: &str, factor: f64) -> String {
    let hex = hex.trim().trim_start_matches('#');
    if hex.len() != 6 {
        return "#E8E8E8".to_string();
    }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64;
    let r = (r + (255.0 - r) * factor) as u8;
    let g = (g + (255.0 - g) * factor) as u8;
    let b = (b + (255.0 - b) * factor) as u8;
    format!("#{:02X}{:02X}{:02X}", r, g, b)
}

#[allow(dead_code)]
fn is_black(color: &str) -> bool {
    let c = color.trim().to_uppercase();
    c == "#000000" || c == "#000" || c == "0"
}

#[allow(dead_code)]
fn is_dark_color(color: &str) -> bool {
    if color.is_empty() || color == "none" {
        return false;
    }
    let c = color.trim().trim_start_matches('#');
    if c.len() == 6 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&c[0..2], 16),
            u8::from_str_radix(&c[2..4], 16),
            u8::from_str_radix(&c[4..6], 16),
        ) {
            let lum = (0.299 * r as f64 + 0.587 * g as f64 + 0.114 * b as f64) / 255.0;
            return lum < 0.4;
        }
    }
    false
}

fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn get_dash_array(pattern: i32, weight: f64) -> String {
    if pattern == 0 {
        return "none".to_string();
    }
    let p = LINE_PATTERNS
        .iter()
        .find(|&&(i, _)| i == pattern)
        .map(|&(_, s)| s)
        .unwrap_or("");
    if p.is_empty() || p == "none" {
        if (2..=23).contains(&pattern) {
            let p = match pattern % 3 {
                0 => "1,2",
                1 => "6,3",
                _ => "6,3,1,3",
            };
            let scale = weight.max(0.5);
            return p
                .split(',')
                .map(|x| format!("{:.1}", x.parse::<f64>().unwrap_or(1.0) * scale))
                .collect::<Vec<_>>()
                .join(",");
        }
        return String::new();
    }
    let scale = weight.max(0.5);
    p.split(',')
        .map(|x| format!("{:.1}", x.parse::<f64>().unwrap_or(1.0) * scale))
        .collect::<Vec<_>>()
        .join(",")
}

/// Merge a shape with its master shape data.
pub fn merge_shape_with_master(
    shape: &mut Shape,
    masters: &HashMap<String, HashMap<String, Shape>>,
    parent_master_id: &str,
) {
    let master_id = if !shape.master.is_empty() {
        &shape.master
    } else if !parent_master_id.is_empty() {
        parent_master_id
    } else {
        return;
    };

    let master_shapes = match masters.get(master_id) {
        Some(ms) => ms,
        None => return,
    };

    let master_sd = if !shape.master_shape.is_empty() {
        master_shapes.get(&shape.master_shape)
    } else {
        master_shapes.values().next()
    };

    let master_sd = match master_sd {
        Some(ms) => ms.clone(),
        None => return,
    };

    // Merge cells
    let mut merged_cells = master_sd.cells.clone();
    for (k, v) in &shape.cells {
        if !v.v.is_empty() || !v.f.is_empty() {
            merged_cells.insert(k.clone(), v.clone());
        }
    }
    shape.cells = merged_cells;

    // Merge geometry
    if shape.geometry.is_empty() && !master_sd.geometry.is_empty() {
        shape.geometry = master_sd.geometry.clone();
        if let Some(cv) = master_sd.cells.get("Width") {
            shape.master_w = cv.as_f64();
        }
        if let Some(cv) = master_sd.cells.get("Height") {
            shape.master_h = cv.as_f64();
        }
    }

    // Merge text
    if shape.text.is_empty()
        && !shape.has_text_elem
        && !master_sd.text.is_empty()
        && shape.shape_type != "Group"
    {
        let txt = &master_sd.text;
        if !matches!(txt.as_str(), "Label" | "Abc" | "Table" | "Entity" | "Class") {
            shape.text = txt.clone();
            if shape.text_parts.is_empty() && !master_sd.text_parts.is_empty() {
                shape.text_parts = master_sd.text_parts.clone();
            }
        }
    }

    // Merge formats
    if shape.char_formats.is_empty() && !master_sd.char_formats.is_empty() {
        shape.char_formats = master_sd.char_formats.clone();
    }
    if shape.para_formats.is_empty() && !master_sd.para_formats.is_empty() {
        shape.para_formats = master_sd.para_formats.clone();
    }
    if shape.foreign_data.is_none() && master_sd.foreign_data.is_some() {
        shape.foreign_data = master_sd.foreign_data.clone();
    }
    if shape.gradient_stops.is_empty() && !master_sd.gradient_stops.is_empty() {
        shape.gradient_stops = master_sd.gradient_stops.clone();
    }
}

/// Compute SVG transform for a shape.
fn compute_transform(shape: &Shape, page_h: f64) -> String {
    let pin_x = shape.cell_f64("PinX") * INCH_TO_PX;
    let pin_y = (page_h - shape.cell_f64("PinY")) * INCH_TO_PX;
    let w = shape.cell_f64("Width");
    let h = shape.cell_f64("Height");

    let lpx_val = shape.cell_val("LocPinX");
    let loc_pin_x = if lpx_val.is_empty() {
        w.abs() * 0.5
    } else {
        lpx_val.parse().unwrap_or(w.abs() * 0.5)
    } * INCH_TO_PX;

    let lpy_val = shape.cell_val("LocPinY");
    let loc_pin_y_raw = if lpy_val.is_empty() {
        h.abs() * 0.5
    } else {
        lpy_val.parse().unwrap_or(h.abs() * 0.5)
    };
    let loc_pin_y = (h.abs() - loc_pin_y_raw) * INCH_TO_PX;

    let angle = shape.cell_f64("Angle");
    let flip_x = shape.cell_val("FlipX") == "1";
    let flip_y = shape.cell_val("FlipY") == "1";

    let mut parts = Vec::new();
    let tx = pin_x - loc_pin_x;
    let ty = pin_y - loc_pin_y;
    parts.push(format!("translate({:.2},{:.2})", tx, ty));

    if angle.abs() > 1e-6 {
        let angle_deg = -angle.to_degrees();
        parts.push(format!(
            "rotate({:.2},{:.2},{:.2})",
            angle_deg, loc_pin_x, loc_pin_y
        ));
    }

    if flip_x || flip_y {
        let sx = if flip_x { -1 } else { 1 };
        let sy = if flip_y { -1 } else { 1 };
        parts.push(format!("translate({:.2},{:.2})", loc_pin_x, loc_pin_y));
        parts.push(format!("scale({},{})", sx, sy));
        parts.push(format!("translate({:.2},{:.2})", -loc_pin_x, -loc_pin_y));
    }

    parts.join(" ")
}

/// Convert geometry to SVG path d attribute.
fn geometry_to_path(geo: &GeomSection, w: f64, h: f64, master_w: f64, master_h: f64) -> String {
    if geo.no_show {
        return String::new();
    }

    let abs_w = if w.abs() > 1e-10 { w.abs() } else { 0.0 };
    let abs_h = if h.abs() > 1e-10 { h.abs() } else { 0.0 };

    let sx = if master_w.abs() > 1e-6 && (master_w.abs() - abs_w).abs() > 1e-6 {
        abs_w / master_w.abs()
    } else {
        1.0
    };
    let sy = if master_h.abs() > 1e-6 && (master_h.abs() - abs_h).abs() > 1e-6 {
        abs_h / master_h.abs()
    } else {
        1.0
    };

    let mut d_parts = Vec::new();
    let mut cx = 0.0_f64;
    let mut cy = 0.0_f64;

    for row in &geo.rows {
        let rt = row.row_type.as_str();
        match rt {
            "MoveTo" => {
                let x = row.cell_f64("X") * sx;
                let y = row.cell_f64("Y") * sy;
                d_parts.push(format!(
                    "M {:.2} {:.2}",
                    x * INCH_TO_PX,
                    (abs_h - y) * INCH_TO_PX
                ));
                cx = x;
                cy = y;
            }
            "RelMoveTo" => {
                let x = row.cell_f64("X") * abs_w;
                let y = row.cell_f64("Y") * abs_h;
                d_parts.push(format!(
                    "M {:.2} {:.2}",
                    x * INCH_TO_PX,
                    (abs_h - y) * INCH_TO_PX
                ));
                cx = x;
                cy = y;
            }
            "LineTo" => {
                let x = row.cell_f64("X") * sx;
                let y = row.cell_f64("Y") * sy;
                d_parts.push(format!(
                    "L {:.2} {:.2}",
                    x * INCH_TO_PX,
                    (abs_h - y) * INCH_TO_PX
                ));
                cx = x;
                cy = y;
            }
            "RelLineTo" => {
                let x = row.cell_f64("X") * abs_w;
                let y = row.cell_f64("Y") * abs_h;
                d_parts.push(format!(
                    "L {:.2} {:.2}",
                    x * INCH_TO_PX,
                    (abs_h - y) * INCH_TO_PX
                ));
                cx = x;
                cy = y;
            }
            "ArcTo" => {
                let x = row.cell_f64("X") * sx;
                let y = row.cell_f64("Y") * sy;
                let a = row.cell_f64("A") * sy;
                append_arc(&mut d_parts, cx, cy, x, y, a, abs_h);
                cx = x;
                cy = y;
            }
            "EllipticalArcTo" => {
                let x = row.cell_f64("X") * sx;
                let y = row.cell_f64("Y") * sy;
                let a = row.cell_f64("A") * sx;
                let b = row.cell_f64("B") * sy;
                let c_angle = row.cell_f64("C");
                let d_ratio = row.cell_f64("D");
                append_elliptical_arc(&mut d_parts, cx, cy, x, y, a, b, d_ratio, c_angle, abs_h);
                cx = x;
                cy = y;
            }
            "RelEllipticalArcTo" => {
                let x = row.cell_f64("X") * abs_w;
                let y = row.cell_f64("Y") * abs_h;
                let a = row.cell_f64("A") * abs_w;
                let b = row.cell_f64("B") * abs_h;
                let c_angle = row.cell_f64("C");
                let d_ratio = row.cell_f64("D");
                append_elliptical_arc(&mut d_parts, cx, cy, x, y, a, b, d_ratio, c_angle, abs_h);
                cx = x;
                cy = y;
            }
            "Ellipse" => {
                let ex = row.cell_f64("X") * sx;
                let ey = row.cell_f64("Y") * sy;
                let ea = row.cell_f64("A") * sx;
                let eb = row.cell_f64("B") * sy;
                let ec = row.cell_f64("C") * sx;
                let ed = row.cell_f64("D") * sy;
                let rx = ((ea - ex).powi(2) + (eb - ey).powi(2)).sqrt().max(0.001) * INCH_TO_PX;
                let ry = ((ec - ex).powi(2) + (ed - ey).powi(2)).sqrt().max(0.001) * INCH_TO_PX;
                let cpx = ex * INCH_TO_PX;
                let cpy = (abs_h - ey) * INCH_TO_PX;
                d_parts.push(format!(
                    "M {:.2} {:.2} A {:.2} {:.2} 0 1 0 {:.2} {:.2} A {:.2} {:.2} 0 1 0 {:.2} {:.2} Z",
                    cpx - rx, cpy, rx, ry, cpx + rx, cpy, rx, ry, cpx - rx, cpy
                ));
            }
            "RelCurveTo" | "RelCubBezTo" => {
                let x = row.cell_f64("X") * abs_w;
                let y = row.cell_f64("Y") * abs_h;
                let a = row.cell_f64("A") * abs_w;
                let b = row.cell_f64("B") * abs_h;
                let c = row.cell_f64("C") * abs_w;
                let d = row.cell_f64("D") * abs_h;
                d_parts.push(format!(
                    "C {:.2} {:.2} {:.2} {:.2} {:.2} {:.2}",
                    a * INCH_TO_PX,
                    (abs_h - b) * INCH_TO_PX,
                    c * INCH_TO_PX,
                    (abs_h - d) * INCH_TO_PX,
                    x * INCH_TO_PX,
                    (abs_h - y) * INCH_TO_PX
                ));
                cx = x;
                cy = y;
            }
            "NURBSTo" => {
                let x = row.cell_f64("X") * sx;
                let y = row.cell_f64("Y") * sy;
                // Simple fallback: line to endpoint
                d_parts.push(format!(
                    "L {:.2} {:.2}",
                    x * INCH_TO_PX,
                    (abs_h - y) * INCH_TO_PX
                ));
                cx = x;
                cy = y;
            }
            "PolylineTo" => {
                let x = row.cell_f64("X") * sx;
                let y = row.cell_f64("Y") * sy;
                d_parts.push(format!(
                    "L {:.2} {:.2}",
                    x * INCH_TO_PX,
                    (abs_h - y) * INCH_TO_PX
                ));
                cx = x;
                cy = y;
            }
            "SplineStart" | "SplineKnot" => {
                let x = row.cell_f64("X") * sx;
                let y = row.cell_f64("Y") * sy;
                if rt == "SplineStart" {
                    d_parts.push(format!(
                        "M {:.2} {:.2}",
                        x * INCH_TO_PX,
                        (abs_h - y) * INCH_TO_PX
                    ));
                } else {
                    d_parts.push(format!(
                        "L {:.2} {:.2}",
                        x * INCH_TO_PX,
                        (abs_h - y) * INCH_TO_PX
                    ));
                }
                cx = x;
                cy = y;
            }
            "InfiniteLine" => {
                let x = row.cell_f64("X") * sx;
                let y = row.cell_f64("Y") * sy;
                let a = row.cell_f64("A") * sx;
                let b = row.cell_f64("B") * sy;
                d_parts.push(format!(
                    "M {:.2} {:.2}",
                    x * INCH_TO_PX,
                    (abs_h - y) * INCH_TO_PX
                ));
                d_parts.push(format!(
                    "L {:.2} {:.2}",
                    a * INCH_TO_PX,
                    (abs_h - b) * INCH_TO_PX
                ));
            }
            _ => {}
        }
    }

    let mut result = d_parts.join(" ");
    if !result.is_empty() && !result.starts_with('M') {
        result = format!("M 0.00 0.00 {}", result);
    }
    result
}

fn append_arc(d_parts: &mut Vec<String>, cx: f64, cy: f64, x: f64, y: f64, bulge: f64, h: f64) {
    if bulge.abs() < 1e-6 {
        d_parts.push(format!(
            "L {:.2} {:.2}",
            x * INCH_TO_PX,
            (h - y) * INCH_TO_PX
        ));
        return;
    }
    let dx = x - cx;
    let dy = y - cy;
    let chord = (dx * dx + dy * dy).sqrt();
    if chord < 1e-10 {
        return;
    }
    let sagitta = bulge.abs();
    let mut radius = (chord * chord / 4.0 + sagitta * sagitta) / (2.0 * sagitta);
    radius = radius.min(chord * 5.0);
    let radius_px = radius * INCH_TO_PX;
    let large_arc = if sagitta > chord / 2.0 { 1 } else { 0 };
    let sweep = if bulge > 0.0 { 0 } else { 1 };
    d_parts.push(format!(
        "A {:.2} {:.2} 0 {} {} {:.2} {:.2}",
        radius_px,
        radius_px,
        large_arc,
        sweep,
        x * INCH_TO_PX,
        (h - y) * INCH_TO_PX
    ));
}

fn append_elliptical_arc(
    d_parts: &mut Vec<String>,
    cx: f64,
    cy: f64,
    x: f64,
    y: f64,
    a: f64,
    b: f64,
    mut d_ratio: f64,
    c_angle: f64,
    h: f64,
) {
    let chord = ((x - cx).powi(2) + (y - cy).powi(2)).sqrt();
    if chord < 1e-10 {
        return;
    }
    let mid_x = (cx + x) / 2.0;
    let mid_y = (cy + y) / 2.0;
    let dist_ctrl = ((a - mid_x).powi(2) + (b - mid_y).powi(2)).sqrt();
    if dist_ctrl < 1e-6 && (d_ratio - 1.0).abs() < 0.01 {
        d_parts.push(format!(
            "L {:.2} {:.2}",
            x * INCH_TO_PX,
            (h - y) * INCH_TO_PX
        ));
        return;
    }
    let angle_deg = if c_angle != 0.0 {
        c_angle.to_degrees()
    } else {
        0.0
    };
    if d_ratio < 0.001 {
        d_ratio = 1.0;
    }
    let half_chord = chord / 2.0;
    let cos_a = if c_angle != 0.0 {
        (-c_angle).cos()
    } else {
        1.0
    };
    let sin_a = if c_angle != 0.0 {
        (-c_angle).sin()
    } else {
        0.0
    };
    let p3_dx = a - mid_x;
    let p3_dy = b - mid_y;
    let p3_lx = cos_a * p3_dx + sin_a * p3_dy;
    let p3_ly = -sin_a * p3_dx + cos_a * p3_dy;
    let sagitta = (p3_lx * p3_lx + p3_ly * p3_ly).sqrt();
    if sagitta < 1e-6 {
        d_parts.push(format!(
            "L {:.2} {:.2}",
            x * INCH_TO_PX,
            (h - y) * INCH_TO_PX
        ));
        return;
    }
    let r = ((half_chord.powi(2) + sagitta.powi(2)) / (2.0 * sagitta)).min(chord * 5.0);
    let rx = r.abs() * INCH_TO_PX;
    let ry = if d_ratio != 0.0 {
        (r / d_ratio).abs() * INCH_TO_PX
    } else {
        rx
    };
    let cross = (x - cx) * (b - cy) - (y - cy) * (a - cx);
    let sweep = if cross < 0.0 { 0 } else { 1 };
    let large_arc = if sagitta > half_chord { 1 } else { 0 };
    d_parts.push(format!(
        "A {:.2} {:.2} {:.1} {} {} {:.2} {:.2}",
        rx.max(0.1),
        ry.max(0.1),
        angle_deg,
        large_arc,
        sweep,
        x * INCH_TO_PX,
        (h - y) * INCH_TO_PX
    ));
}

/// Render all shapes on a page to SVG.
pub fn shapes_to_svg(
    shapes: &[Shape],
    page_w: f64,
    page_h: f64,
    masters: &HashMap<String, HashMap<String, Shape>>,
    _connects: &[Connect],
    media: &HashMap<String, Vec<u8>>,
    page_rels: &HashMap<String, String>,
    bg_shapes: Option<&[Shape]>,
    theme_colors: &HashMap<String, String>,
    layers: &HashMap<String, LayerDef>,
) -> String {
    let page_w_px = page_w * INCH_TO_PX;
    let page_h_px = page_h * INCH_TO_PX;

    // Compute content bounding box
    let mut vb_x = 0.0_f64;
    let mut vb_y = 0.0_f64;
    let mut vb_w = page_w_px;
    let mut vb_h = page_h_px;

    let all_shapes: Vec<&Shape> = shapes
        .iter()
        .chain(bg_shapes.unwrap_or(&[]).iter())
        .collect();

    if !all_shapes.is_empty() {
        let mut min_x = f64::MAX;
        let mut min_y = f64::MAX;
        let mut max_x = f64::MIN;
        let mut max_y = f64::MIN;

        for s in &all_shapes {
            let px = s.cell_f64("PinX") * INCH_TO_PX;
            let py = (page_h - s.cell_f64("PinY")) * INCH_TO_PX;
            let sw = s.cell_f64("Width").abs() * INCH_TO_PX;
            let sh = s.cell_f64("Height").abs() * INCH_TO_PX;
            if px > 0.0 || py > 0.0 {
                min_x = min_x.min(px - sw / 2.0);
                min_y = min_y.min(py - sh / 2.0);
                max_x = max_x.max(px + sw / 2.0);
                max_y = max_y.max(py + sh / 2.0);
            }
        }

        if min_x < f64::MAX {
            let content_w = max_x - min_x;
            let content_h = max_y - min_y;
            let pad_x = (content_w * 0.08).max(50.0);
            let pad_y = (content_h * 0.08).max(50.0);
            vb_x = min_x.min(0.0) - pad_x;
            vb_y = min_y.min(0.0) - pad_y;
            vb_w = (content_w + 2.0 * pad_x).max(page_w_px * 0.5);
            vb_h = (content_h + 2.0 * pad_y).max(page_h_px * 0.5);
            vb_w = vb_w.max(max_x + pad_x - vb_x);
            vb_h = vb_h.max(max_y + pad_y - vb_y);
        }
    }

    let max_svg_px = 4000.0;
    let mut display_w = vb_w;
    let mut display_h = vb_h;
    if vb_w.max(vb_h) > max_svg_px {
        let scale = max_svg_px / vb_w.max(vb_h);
        display_w = vb_w * scale;
        display_h = vb_h * scale;
    }

    let mut svg = Vec::new();
    svg.push(r#"<?xml version="1.0" encoding="UTF-8"?>"#.to_string());
    svg.push(format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{:.0}" height="{:.0}" viewBox="{:.0} {:.0} {:.0} {:.0}">"#,
        display_w, display_h, vb_x, vb_y, vb_w, vb_h
    ));
    svg.push(format!(
        r#"<rect x="{:.0}" y="{:.0}" width="{:.0}" height="{:.0}" fill="white"/>"#,
        vb_x, vb_y, vb_w, vb_h
    ));

    let mut used_markers: HashSet<String> = HashSet::new();
    let mut gradients: Vec<GradientDef> = Vec::new();
    let mut fill_patterns: Vec<FillPatternDef> = Vec::new();
    let mut shadow_defs: Vec<String> = Vec::new();
    let mut text_layer: Vec<String> = Vec::new();

    // Render background shapes
    if let Some(bg) = bg_shapes {
        svg.push("<!-- Background page -->".to_string());
        for s in bg {
            let mut shape = s.clone();
            merge_shape_with_master(&mut shape, masters, "");
            let elements = render_shape_svg(
                &shape,
                page_h,
                masters,
                "",
                0,
                media,
                page_rels,
                &mut used_markers,
                theme_colors,
                layers,
                &mut gradients,
                &mut fill_patterns,
                &mut shadow_defs,
                &mut text_layer,
            );
            svg.extend(elements);
        }
    }

    // Sort: containers first
    let mut sorted_shapes: Vec<&Shape> = shapes.iter().collect();
    sorted_shapes.sort_by_key(|s| {
        let user = &s.user;
        let st = user
            .get("msvStructureType")
            .and_then(|m| m.get("Value"))
            .map(|v| v.as_str())
            .unwrap_or("");
        let name = s.name_u.to_lowercase();
        if st == "Container" || name.contains("container") || name.contains("swimlane") {
            0
        } else {
            1
        }
    });

    // Render shapes
    for s in sorted_shapes {
        let mut shape = s.clone();
        merge_shape_with_master(&mut shape, masters, "");
        let elements = render_shape_svg(
            &shape,
            page_h,
            masters,
            "",
            0,
            media,
            page_rels,
            &mut used_markers,
            theme_colors,
            layers,
            &mut gradients,
            &mut fill_patterns,
            &mut shadow_defs,
            &mut text_layer,
        );
        svg.extend(elements);
    }

    // Text layer on top
    if !text_layer.is_empty() {
        svg.push("<!-- Text layer -->".to_string());
        svg.extend(text_layer);
    }

    svg.push("</svg>".to_string());

    // Insert defs block after background rect
    let mut defs = Vec::new();
    if !used_markers.is_empty()
        || !gradients.is_empty()
        || !fill_patterns.is_empty()
        || !shadow_defs.is_empty()
    {
        defs.push("<defs>".to_string());

        // Arrow markers
        for marker_id in &used_markers {
            let parts: Vec<&str> = marker_id.split('_').collect();
            let direction = parts.get(1).unwrap_or(&"end");
            let size_idx: i32 = parts.get(2).and_then(|s| s.parse().ok()).unwrap_or(3);
            let color = parts
                .get(3)
                .map(|c| format!("#{}", c))
                .unwrap_or_else(|| "#333333".to_string());
            let scale = arrow_size(size_idx);
            let mw = 10.0 * scale;
            let mh = 7.0 * scale;
            if *direction == "start" {
                defs.push(format!(
                    r#"<marker id="{}" markerWidth="{:.1}" markerHeight="{:.1}" refX="0" refY="{:.1}" orient="auto" markerUnits="userSpaceOnUse"><polygon points="{:.1} 0, 0 {:.1}, {:.1} {:.1}" fill="{}"/></marker>"#,
                    marker_id, mw, mh, mh / 2.0, mw, mh / 2.0, mw, mh, color
                ));
            } else {
                defs.push(format!(
                    r#"<marker id="{}" markerWidth="{:.1}" markerHeight="{:.1}" refX="{:.1}" refY="{:.1}" orient="auto" markerUnits="userSpaceOnUse"><polygon points="0 0, {:.1} {:.1}, 0 {:.1}" fill="{}"/></marker>"#,
                    marker_id, mw, mh, mw, mh / 2.0, mw, mh / 2.0, mh, color
                ));
            }
        }

        // Fill patterns
        for pat in &fill_patterns {
            let spacing = 6;
            let sw = 1.0;
            let pt = pat.pattern_type;
            if (2..=5).contains(&pt) {
                let line = match pt {
                    2 => format!(
                        r#"<line x1="0" y1="3" x2="6" y2="3" stroke="{}" stroke-width="{}"/>"#,
                        pat.fg, sw
                    ),
                    3 => format!(
                        r#"<line x1="3" y1="0" x2="3" y2="6" stroke="{}" stroke-width="{}"/>"#,
                        pat.fg, sw
                    ),
                    4 => format!(
                        r#"<line x1="0" y1="6" x2="6" y2="0" stroke="{}" stroke-width="{}"/>"#,
                        pat.fg, sw
                    ),
                    _ => format!(
                        r#"<line x1="0" y1="0" x2="6" y2="6" stroke="{}" stroke-width="{}"/>"#,
                        pat.fg, sw
                    ),
                };
                defs.push(format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}"><rect width="{}" height="{}" fill="{}"/>{}</pattern>"#,
                    pat.id, spacing, spacing, spacing, spacing, pat.bg, line
                ));
            } else if (6..=9).contains(&pt) {
                defs.push(format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}"><rect width="{}" height="{}" fill="{}"/><line x1="0" y1="3" x2="6" y2="3" stroke="{}" stroke-width="{}"/><line x1="3" y1="0" x2="3" y2="6" stroke="{}" stroke-width="{}"/></pattern>"#,
                    pat.id, spacing, spacing, spacing, spacing, pat.bg, pat.fg, sw, pat.fg, sw
                ));
            } else {
                defs.push(format!(
                    r#"<pattern id="{}" patternUnits="userSpaceOnUse" width="{}" height="{}"><rect width="{}" height="{}" fill="{}"/><circle cx="3" cy="3" r="0.8" fill="{}"/></pattern>"#,
                    pat.id, spacing, spacing, spacing, spacing, pat.bg, pat.fg
                ));
            }
        }

        // Gradients
        for grad in &gradients {
            let stops = if !grad.stops.is_empty() {
                grad.stops
                    .iter()
                    .map(|s| {
                        format!(
                            r#"<stop offset="{}%" stop-color="{}"/>"#,
                            s.position, s.color
                        )
                    })
                    .collect::<Vec<_>>()
                    .join("")
            } else {
                let mut s = vec![format!(
                    r#"<stop offset="0%" stop-color="{}"/>"#,
                    grad.start
                )];
                if let Some(mid) = &grad.mid {
                    s.push(format!(r#"<stop offset="50%" stop-color="{}"/>"#, mid));
                }
                s.push(format!(
                    r#"<stop offset="100%" stop-color="{}"/>"#,
                    grad.end
                ));
                s.join("")
            };

            if grad.radial {
                defs.push(format!(
                    r#"<radialGradient id="{}" cx="{}%" cy="{}%" r="{}%" fx="{}%" fy="{}%">{}</radialGradient>"#,
                    grad.id, grad.cx, grad.cy, grad.r, grad.fx, grad.fy, stops
                ));
            } else {
                let rad = grad.dir.to_radians();
                let x1 = 50.0 - 50.0 * rad.cos();
                let y1 = 50.0 + 50.0 * rad.sin();
                let x2 = 50.0 + 50.0 * rad.cos();
                let y2 = 50.0 - 50.0 * rad.sin();
                defs.push(format!(
                    r#"<linearGradient id="{}" x1="{:.1}%" y1="{:.1}%" x2="{:.1}%" y2="{:.1}%">{}</linearGradient>"#,
                    grad.id, x1, y1, x2, y2, stops
                ));
            }
        }

        // Shadow filters
        for sdef in &shadow_defs {
            defs.push(sdef.clone());
        }

        defs.push("</defs>".to_string());
    }

    // Insert defs after line 3 (after background rect)
    if !defs.is_empty() {
        for (j, line) in defs.into_iter().enumerate() {
            svg.insert(3 + j, line);
        }
    }

    svg.join("\n")
}

/// Render a single shape as SVG elements.
fn render_shape_svg(
    shape: &Shape,
    page_h: f64,
    masters: &HashMap<String, HashMap<String, Shape>>,
    parent_master_id: &str,
    depth: usize,
    media: &HashMap<String, Vec<u8>>,
    page_rels: &HashMap<String, String>,
    used_markers: &mut HashSet<String>,
    theme_colors: &HashMap<String, String>,
    layers: &HashMap<String, LayerDef>,
    gradients: &mut Vec<GradientDef>,
    fill_patterns: &mut Vec<FillPatternDef>,
    shadow_defs: &mut Vec<String>,
    text_layer: &mut Vec<String>,
) -> Vec<String> {
    let mut lines = Vec::new();

    // Skip invisible shapes
    if shape.cell_val("Visible") == "0" {
        return lines;
    }

    // Layer visibility check
    let layer_member = shape.cell_val("LayerMember");
    if !layer_member.is_empty() && !layers.is_empty() {
        let layer_ids: Vec<&str> = layer_member.split(';').collect();
        let all_hidden = layer_ids
            .iter()
            .all(|lid| layers.get(lid.trim()).map(|l| !l.visible).unwrap_or(false));
        if all_hidden {
            return lines;
        }
    }

    let w_inch = shape.cell_f64("Width");
    let h_inch = shape.cell_f64("Height");
    let w_px = w_inch.abs() * INCH_TO_PX;
    let h_px = h_inch.abs() * INCH_TO_PX;

    // Line style
    let mut line_weight = shape.cell_f64_or("LineWeight", 0.01) * INCH_TO_PX;
    if line_weight < 0.5 {
        line_weight = 1.5;
    } else if line_weight > 20.0 {
        line_weight = 20.0;
    }

    let line_color = {
        let c = resolve_color(shape.cell_val("LineColor"), theme_colors);
        if c.is_empty() {
            "#333333".to_string()
        } else {
            c
        }
    };

    let fill_foregnd = resolve_color(shape.cell_val("FillForegnd"), theme_colors);
    let fill_bkgnd = resolve_color(shape.cell_val("FillBkgnd"), theme_colors);

    let fill_pat_int: i32 = shape.cell_val("FillPattern").parse().unwrap_or(1);
    let line_pattern: i32 = shape.cell_val("LinePattern").parse().unwrap_or(1);

    // Determine fill color
    let fill = if fill_pat_int == 0 {
        "none".to_string()
    } else if fill_pat_int == 1 {
        if !fill_foregnd.is_empty() {
            fill_foregnd.clone()
        } else if !fill_bkgnd.is_empty() {
            fill_bkgnd.clone()
        } else {
            "none".to_string()
        }
    } else if !fill_foregnd.is_empty() {
        fill_foregnd.clone()
    } else if !fill_bkgnd.is_empty() {
        fill_bkgnd.clone()
    } else {
        "none".to_string()
    };

    let stroke = if line_pattern != 0 {
        &line_color
    } else {
        "none"
    };
    let dash_array = get_dash_array(line_pattern, line_weight);

    // Fill opacity
    let fill_trans = shape.cell_f64("FillForegndTrans");
    let fill_opacity = if fill_trans > 0.0 && fill_trans <= 1.0 {
        1.0 - fill_trans
    } else if fill_trans > 1.0 {
        1.0 - fill_trans / 100.0
    } else {
        1.0
    };

    // Check for 1D connector
    let begin_x = shape.cell_val("BeginX");
    let end_x = shape.cell_val("EndX");
    let is_1d = !begin_x.is_empty() && !end_x.is_empty();
    let is_1d_group = (shape.shape_type == "Group" || !shape.sub_shapes.is_empty()) && is_1d;

    // Group rendering
    if (shape.shape_type == "Group" || !shape.sub_shapes.is_empty()) && !is_1d_group {
        let transform = compute_transform(shape, page_h);
        let group_master_id = if !shape.master.is_empty() {
            &shape.master
        } else {
            parent_master_id
        };
        let group_h = h_inch;

        lines.push(format!(r#"<g transform="{}">"#, transform));

        // Render group's own geometry
        for geo in &shape.geometry {
            let path_d = geometry_to_path(geo, w_inch, h_inch, shape.master_w, shape.master_h);
            if path_d.is_empty() {
                continue;
            }
            let geo_fill = if geo.no_fill { "none" } else { &fill };
            let geo_stroke = if geo.no_line { "none" } else { stroke };
            lines.push(format!(
                r#"<path d="{}" fill="{}" stroke="{}" stroke-width="{:.2}"/>"#,
                path_d, geo_fill, geo_stroke, line_weight
            ));
        }

        // Render sub-shapes
        for sub in &shape.sub_shapes {
            let mut sub_shape = sub.clone();
            merge_shape_with_master(&mut sub_shape, masters, group_master_id);
            let sub_elements = render_shape_svg(
                &sub_shape,
                group_h,
                masters,
                group_master_id,
                depth + 1,
                media,
                page_rels,
                used_markers,
                theme_colors,
                layers,
                gradients,
                fill_patterns,
                shadow_defs,
                text_layer,
            );
            lines.extend(sub_elements);
        }

        lines.push("</g>".to_string());

        // Group text
        if !shape.text.is_empty() && depth == 0 {
            append_text_svg(text_layer, shape, page_h, w_px, h_px, theme_colors);
        } else if !shape.text.is_empty() {
            append_text_svg(&mut lines, shape, page_h, w_px, h_px, theme_colors);
        }

        return lines;
    }

    let transform = compute_transform(shape, page_h);

    // 1D connector rendering
    if is_1d {
        let bx = begin_x.parse::<f64>().unwrap_or(0.0) * INCH_TO_PX;
        let by = (page_h - shape.cell_val("BeginY").parse::<f64>().unwrap_or(0.0)) * INCH_TO_PX;
        let ex_px = end_x.parse::<f64>().unwrap_or(0.0) * INCH_TO_PX;
        let ey_px = (page_h - shape.cell_val("EndY").parse::<f64>().unwrap_or(0.0)) * INCH_TO_PX;

        // Arrow markers
        let end_arrow: i32 = shape.cell_val("EndArrow").parse().unwrap_or(0);
        let begin_arrow: i32 = shape.cell_val("BeginArrow").parse().unwrap_or(0);
        let end_arrow_size: i32 = shape.cell_val("EndArrowSize").parse().unwrap_or(2);
        let begin_arrow_size: i32 = shape.cell_val("BeginArrowSize").parse().unwrap_or(2);
        let marker_color = stroke.trim_start_matches('#');
        let mut marker_attrs = String::new();
        if begin_arrow > 0 {
            let mid = format!("arrow_start_{}_{}", begin_arrow_size, marker_color);
            used_markers.insert(mid.clone());
            marker_attrs.push_str(&format!(r#" marker-start="url(#{})""#, mid));
        }
        if end_arrow > 0 {
            let mid = format!("arrow_end_{}_{}", end_arrow_size, marker_color);
            used_markers.insert(mid.clone());
            marker_attrs.push_str(&format!(r#" marker-end="url(#{})""#, mid));
        }

        let dash_attr = if !dash_array.is_empty() {
            format!(r#" stroke-dasharray="{}""#, dash_array)
        } else {
            String::new()
        };

        lines.push(format!(
            r#"<line x1="{:.2}" y1="{:.2}" x2="{:.2}" y2="{:.2}" stroke="{}" stroke-width="{:.2}"{}{}/>
"#,
            bx, by, ex_px, ey_px, stroke, line_weight.max(1.5), dash_attr, marker_attrs
        ));
    } else if !shape.geometry.is_empty() {
        // 2D shape with geometry
        let _rounding = shape.cell_f64("Rounding") * INCH_TO_PX;

        for geo in &shape.geometry {
            let path_d = geometry_to_path(geo, w_inch, h_inch, shape.master_w, shape.master_h);
            if path_d.is_empty() {
                continue;
            }
            let geo_fill = if geo.no_fill { "none" } else { &fill };
            let geo_stroke = if geo.no_line { "none" } else { stroke };

            let mut style = format!(
                r#"fill="{}" stroke="{}" stroke-width="{:.2}""#,
                geo_fill, geo_stroke, line_weight
            );
            if fill_opacity < 0.99 && geo_fill != "none" {
                style.push_str(&format!(r#" fill-opacity="{:.2}""#, fill_opacity));
            }
            if !dash_array.is_empty() {
                style.push_str(&format!(r#" stroke-dasharray="{}""#, dash_array));
            }
            lines.push(format!(
                r#"<path d="{}" {} transform="{}"/>"#,
                path_d, style, transform
            ));
        }
    } else if w_px > 0.0 && h_px > 0.0 && (fill != "none" || !shape.text.is_empty()) {
        // Fallback rectangle
        let fallback_fill = if fill != "none" { &fill } else { "#FAFAFA" };
        let fallback_stroke = if stroke != "none" { stroke } else { "#CCCCCC" };
        lines.push(format!(
            r#"<rect x="0" y="0" width="{:.2}" height="{:.2}" fill="{}" stroke="{}" stroke-width="{:.2}" rx="4" transform="{}"/>"#,
            w_px, h_px, fallback_fill, fallback_stroke, line_weight.max(0.75), transform
        ));
    }

    // Text rendering
    if !shape.text.is_empty() {
        if depth == 0 {
            append_text_svg(text_layer, shape, page_h, w_px, h_px, theme_colors);
        } else {
            append_text_svg(&mut lines, shape, page_h, w_px, h_px, theme_colors);
        }
    }

    lines
}

/// Append SVG text elements for a shape.
fn append_text_svg(
    lines: &mut Vec<String>,
    shape: &Shape,
    page_h: f64,
    _w_px: f64,
    _h_px: f64,
    theme_colors: &HashMap<String, String>,
) {
    let text = &shape.text;
    if text.is_empty() {
        return;
    }

    let pin_x = shape.cell_f64("PinX") * INCH_TO_PX;
    let pin_y = (page_h - shape.cell_f64("PinY")) * INCH_TO_PX;

    let char_fmt = shape.char_formats.get("0").cloned().unwrap_or_default();
    let mut font_size = char_fmt.size.parse::<f64>().unwrap_or(0.1111) * INCH_TO_PX;
    if font_size < 6.0 {
        font_size = 8.0;
    } else if font_size > 72.0 {
        font_size = 72.0;
    }

    let text_color = {
        let c = resolve_color(&char_fmt.color, theme_colors);
        if c.is_empty() {
            "#000000".to_string()
        } else {
            c
        }
    };

    let font_family = if char_fmt.font.is_empty() || char_fmt.font == "Themed" {
        "Noto Sans, sans-serif".to_string()
    } else {
        format!("{}, Noto Sans, sans-serif", char_fmt.font)
    };

    let para_fmt = shape.para_formats.get("0").cloned().unwrap_or_default();
    let halign: i32 = para_fmt.horiz_align.parse().unwrap_or(1);
    let text_anchor = match halign {
        0 => "start",
        2 => "end",
        _ => "middle",
    };

    let style_bits: i32 = char_fmt.style.parse().unwrap_or(0);
    let fw = if style_bits & 1 != 0 {
        r#" font-weight="bold""#
    } else {
        ""
    };
    let fs = if style_bits & 2 != 0 {
        r#" font-style="italic""#
    } else {
        ""
    };

    let tx = pin_x;
    let ty = pin_y;

    let text_lines: Vec<&str> = text.split('\n').collect();
    let total_height = text_lines.len() as f64 * font_size * 1.2;
    let start_y = ty - total_height / 2.0 + font_size * 0.6;

    for (j, tline) in text_lines.iter().enumerate() {
        let trimmed = tline.trim();
        if trimmed.is_empty() {
            continue;
        }
        let escaped = escape_xml(trimmed);
        let ly = start_y + j as f64 * font_size * 1.2;
        lines.push(format!(
            r#"<text x="{:.2}" y="{:.2}" text-anchor="{}" font-family="{}" font-size="{:.1}" fill="{}"{}{}>
{}</text>"#,
            tx, ly, text_anchor, font_family, font_size, text_color, fw, fs, escaped
        ));
    }
}
