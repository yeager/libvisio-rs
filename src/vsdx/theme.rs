//! Theme color resolution from DrawingML theme XML.

use std::collections::HashMap;

/// Parse theme colors from visio/theme/theme1.xml inside the ZIP.
pub fn parse_theme(zip: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> HashMap<String, String> {
    let mut theme_colors = HashMap::new();
    let dml_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";

    for theme_file in &["visio/theme/theme1.xml", "visio/theme/theme2.xml"] {
        let xml_data = match zip.by_name(theme_file) {
            Ok(mut f) => {
                let mut buf = Vec::new();
                std::io::Read::read_to_end(&mut f, &mut buf).ok();
                buf
            }
            Err(_) => continue,
        };
        let xml_str = match std::str::from_utf8(&xml_data) {
            Ok(s) => s,
            Err(_) => continue,
        };
        let doc = match roxmltree::Document::parse(xml_str) {
            Ok(d) => d,
            Err(_) => continue,
        };

        let color_names = [
            "dk1", "lt1", "dk2", "lt2",
            "accent1", "accent2", "accent3", "accent4",
            "accent5", "accent6", "hlink", "folHlink",
        ];

        // Find clrScheme element
        for node in doc.descendants() {
            if node.tag_name().name() == "clrScheme" {
                for cname in &color_names {
                    for child in node.children() {
                        if child.tag_name().name() == *cname {
                            if let Some(color) = extract_color(&child, dml_ns) {
                                theme_colors.insert(cname.to_string(), color);
                            }
                            break;
                        }
                    }
                }
                break;
            }
        }

        if !theme_colors.is_empty() {
            break;
        }
    }

    // Build numeric index mapping
    let idx_map = [
        (0, "dk1"), (1, "lt1"), (2, "dk2"), (3, "lt2"),
        (4, "accent1"), (5, "accent2"), (6, "accent3"), (7, "accent4"),
        (8, "accent5"), (9, "accent6"), (10, "hlink"), (11, "folHlink"),
    ];
    for (idx, name) in &idx_map {
        if let Some(color) = theme_colors.get(*name) {
            theme_colors.insert(idx.to_string(), color.clone());
        }
    }

    theme_colors
}

fn extract_color(elem: &roxmltree::Node, _dml_ns: &str) -> Option<String> {
    for child in elem.children() {
        let tag = child.tag_name().name();
        if tag == "srgbClr" {
            if let Some(val) = child.attribute("val") {
                let base = format!("#{}", val);
                return Some(apply_color_transforms(&child, &base));
            }
        } else if tag == "sysClr" {
            let val = child.attribute("lastClr")
                .or_else(|| child.attribute("val"))
                .unwrap_or("");
            if val.len() == 6 {
                let base = format!("#{}", val);
                return Some(apply_color_transforms(&child, &base));
            }
        }
    }
    None
}

fn apply_color_transforms(elem: &roxmltree::Node, base_color: &str) -> String {
    if base_color.len() != 7 || !base_color.starts_with('#') {
        return base_color.to_string();
    }
    let r = u8::from_str_radix(&base_color[1..3], 16).unwrap_or(0) as f64;
    let g = u8::from_str_radix(&base_color[3..5], 16).unwrap_or(0) as f64;
    let b = u8::from_str_radix(&base_color[5..7], 16).unwrap_or(0) as f64;
    let mut r = r;
    let mut g = g;
    let mut b = b;

    for child in elem.children() {
        let tag = child.tag_name().name();
        let val: f64 = child.attribute("val")
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(0.0);
        let pct = val / 100000.0;

        match tag {
            "tint" => {
                r += (255.0 - r) * pct;
                g += (255.0 - g) * pct;
                b += (255.0 - b) * pct;
            }
            "shade" => {
                r *= pct;
                g *= pct;
                b *= pct;
            }
            "lumMod" => {
                r = (r * pct).min(255.0);
                g = (g * pct).min(255.0);
                b = (b * pct).min(255.0);
            }
            "lumOff" => {
                let off = 255.0 * pct;
                r = (r + off).clamp(0.0, 255.0);
                g = (g + off).clamp(0.0, 255.0);
                b = (b + off).clamp(0.0, 255.0);
            }
            _ => {}
        }
    }

    format!("#{:02X}{:02X}{:02X}", r.clamp(0.0, 255.0) as u8, g.clamp(0.0, 255.0) as u8, b.clamp(0.0, 255.0) as u8)
}
