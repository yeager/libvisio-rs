//! .vsdx (ZIP + XML) parser.
//!
//! Parses the Open Packaging Convention ZIP file containing Visio XML.

use std::collections::HashMap;
use std::io::{Cursor, Read};

use crate::error::{Result, VisioError};
use crate::model::*;
use crate::vsdx::image;
use crate::vsdx::theme;

const VNS: &str = "http://schemas.microsoft.com/office/visio/2012/main";
const RNS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn is_visio_tag(node: &roxmltree::Node, name: &str) -> bool {
    node.tag_name().name() == name
        && (node.tag_name().namespace().is_none() || node.tag_name().namespace() == Some(VNS))
}

fn find_child<'a>(
    node: &'a roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    node.children().find(|c| is_visio_tag(c, name))
}

fn cell_val(node: &roxmltree::Node, cell_name: &str) -> String {
    for child in node.children() {
        if is_visio_tag(&child, "Cell") && child.attribute("N") == Some(cell_name) {
            return child.attribute("V").unwrap_or("").to_string();
        }
    }
    String::new()
}

/// Parse a complete .vsdx file from bytes.
pub fn parse_vsdx(data: &[u8]) -> Result<Document> {
    let cursor = Cursor::new(data);
    let mut zip = zip::ZipArchive::new(cursor).map_err(|e| VisioError::Zip(e))?;

    let mut doc = Document::default();

    // Parse theme colors
    doc.theme_colors = theme::parse_theme(&mut zip);

    // Parse media
    doc.media = image::extract_media(&mut zip);

    // Parse master shapes
    doc.masters = parse_master_shapes(&mut zip);

    // Parse stylesheets
    doc.stylesheets = parse_stylesheets(&mut zip);

    // Parse background page map
    doc.background_map = parse_background_pages(&mut zip);

    // Parse page names
    let page_names = parse_page_names(&mut zip);

    // Parse page dimensions from pages.xml
    let all_dims = parse_all_page_dimensions(&mut zip);

    // Get page files
    let page_files = get_page_files(&mut zip);

    // Parse master rels
    let master_rels = parse_master_rels(&mut zip);

    // Pre-parse all pages
    let mut page_cache: HashMap<
        usize,
        (
            Vec<Shape>,
            Vec<Connect>,
            HashMap<String, String>,
            HashMap<String, LayerDef>,
        ),
    > = HashMap::new();

    for (i, page_file) in page_files.iter().enumerate() {
        let page_xml = match read_zip_file(&mut zip, page_file) {
            Some(data) => data,
            None => continue,
        };
        let xml_str = match std::str::from_utf8(&page_xml) {
            Ok(s) => s,
            Err(_) => continue,
        };
        let xml_doc = match roxmltree::Document::parse(xml_str) {
            Ok(d) => d,
            Err(_) => continue,
        };

        let shapes = parse_page_shapes(&xml_doc);
        let connects = parse_connects(&xml_doc);
        let layers = parse_layers(&xml_doc);
        let page_rels = image::parse_rels(&mut zip, page_file);

        page_cache.insert(i, (shapes, connects, page_rels, layers));
    }

    // Build pages
    for (i, _page_file) in page_files.iter().enumerate() {
        let (shapes, connects, page_rels, layers) = match page_cache.remove(&i) {
            Some(data) => data,
            None => continue,
        };

        if shapes.is_empty() {
            continue;
        }

        let (page_w, page_h) = if i < all_dims.len() {
            all_dims[i]
        } else {
            (8.5, 11.0)
        };

        let name = page_names
            .get(i)
            .cloned()
            .unwrap_or_else(|| format!("Page {}", i + 1));

        let mut all_rels = master_rels.clone();
        all_rels.extend(page_rels);

        let page = Page {
            name,
            index: i,
            width: page_w,
            height: page_h,
            shapes,
            connects,
            layers,
            background: false,
        };

        doc.pages.push(page);
    }

    Ok(doc)
}

fn read_zip_file(zip: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str) -> Option<Vec<u8>> {
    let mut f = zip.by_name(name).ok()?;
    let mut buf = Vec::new();
    f.read_to_end(&mut buf).ok()?;
    Some(buf)
}

fn get_page_files(zip: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let mut page_files: Vec<String> = (0..zip.len())
        .filter_map(|i| {
            let f = zip.by_index(i).ok()?;
            let name = f.name().to_string();
            if name.starts_with("visio/pages/page")
                && name.ends_with(".xml")
                && !name.ends_with("pages.xml")
            {
                Some(name)
            } else {
                None
            }
        })
        .collect();
    page_files.sort();
    page_files
}

fn parse_page_names(zip: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let mut names = Vec::new();
    let data = match read_zip_file(zip, "visio/pages/pages.xml") {
        Some(d) => d,
        None => return names,
    };
    let xml_str = match std::str::from_utf8(&data) {
        Ok(s) => s,
        Err(_) => return names,
    };
    let doc = match roxmltree::Document::parse(xml_str) {
        Ok(d) => d,
        Err(_) => return names,
    };
    for node in doc.descendants() {
        if is_visio_tag(&node, "Page") {
            names.push(node.attribute("Name").unwrap_or("").to_string());
        }
    }
    names
}

fn parse_all_page_dimensions(zip: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<(f64, f64)> {
    let mut dims = Vec::new();
    let data = match read_zip_file(zip, "visio/pages/pages.xml") {
        Some(d) => d,
        None => return dims,
    };
    let xml_str = match std::str::from_utf8(&data) {
        Ok(s) => s,
        Err(_) => return dims,
    };
    let doc = match roxmltree::Document::parse(xml_str) {
        Ok(d) => d,
        Err(_) => return dims,
    };
    for node in doc.descendants() {
        if is_visio_tag(&node, "Page") {
            let mut pw = 8.5;
            let mut ph = 11.0;
            if let Some(ps) = find_child(&node, "PageSheet") {
                for cell in ps.children() {
                    if is_visio_tag(&cell, "Cell") {
                        match cell.attribute("N") {
                            Some("PageWidth") => {
                                pw = cell
                                    .attribute("V")
                                    .and_then(|v| v.parse().ok())
                                    .unwrap_or(8.5);
                            }
                            Some("PageHeight") => {
                                ph = cell
                                    .attribute("V")
                                    .and_then(|v| v.parse().ok())
                                    .unwrap_or(11.0);
                            }
                            _ => {}
                        }
                    }
                }
            }
            dims.push((pw, ph));
        }
    }
    dims
}

fn parse_background_pages(zip: &mut zip::ZipArchive<Cursor<&[u8]>>) -> HashMap<usize, usize> {
    let mut bg_map = HashMap::new();
    let data = match read_zip_file(zip, "visio/pages/pages.xml") {
        Some(d) => d,
        None => return bg_map,
    };
    let xml_str = match std::str::from_utf8(&data) {
        Ok(s) => s,
        Err(_) => return bg_map,
    };
    let doc = match roxmltree::Document::parse(xml_str) {
        Ok(d) => d,
        Err(_) => return bg_map,
    };

    let mut page_id_to_idx: HashMap<String, usize> = HashMap::new();
    let mut pages_data: Vec<(usize, roxmltree::Node)> = Vec::new();

    for (i, node) in doc
        .descendants()
        .filter(|n| is_visio_tag(n, "Page"))
        .enumerate()
    {
        if let Some(pid) = node.attribute("ID") {
            page_id_to_idx.insert(pid.to_string(), i);
        }
        pages_data.push((i, node));
    }

    for (i, node) in &pages_data {
        if let Some(ps) = find_child(node, "PageSheet") {
            for cell in ps.children() {
                if is_visio_tag(&cell, "Cell") && cell.attribute("N") == Some("BackPage") {
                    if let Some(back_id) = cell.attribute("V") {
                        if let Some(&bg_idx) = page_id_to_idx.get(back_id) {
                            bg_map.insert(*i, bg_idx);
                        }
                    }
                }
            }
        }
    }
    bg_map
}

fn parse_stylesheets(zip: &mut zip::ZipArchive<Cursor<&[u8]>>) -> HashMap<String, StyleSheet> {
    let mut styles = HashMap::new();
    let data = match read_zip_file(zip, "visio/document.xml") {
        Some(d) => d,
        None => return styles,
    };
    let xml_str = match std::str::from_utf8(&data) {
        Ok(s) => s,
        Err(_) => return styles,
    };
    let doc = match roxmltree::Document::parse(xml_str) {
        Ok(d) => d,
        Err(_) => return styles,
    };

    for node in doc.descendants() {
        if is_visio_tag(&node, "StyleSheet") {
            let sid = node.attribute("ID").unwrap_or("").to_string();
            if sid.is_empty() {
                continue;
            }
            let mut ss = StyleSheet::default();
            ss.line_style = node.attribute("LineStyle").unwrap_or("").to_string();
            ss.fill_style = node.attribute("FillStyle").unwrap_or("").to_string();
            ss.text_style = node.attribute("TextStyle").unwrap_or("").to_string();
            for cell in node.children() {
                if is_visio_tag(&cell, "Cell") {
                    let n = cell.attribute("N").unwrap_or("");
                    let v = cell.attribute("V").unwrap_or("");
                    let f = cell.attribute("F").unwrap_or("");
                    ss.cells.insert(n.to_string(), CellValue::new(v, f));
                }
            }
            styles.insert(sid, ss);
        }
    }
    styles
}

fn parse_master_rels(zip: &mut zip::ZipArchive<Cursor<&[u8]>>) -> HashMap<String, String> {
    let mut rels = HashMap::new();
    let names: Vec<String> = (0..zip.len())
        .filter_map(|i| {
            let f = zip.by_index(i).ok()?;
            let n = f.name().to_string();
            if n.starts_with("visio/masters/_rels/master") && n.ends_with(".xml.rels") {
                Some(n)
            } else {
                None
            }
        })
        .collect();

    for name in names {
        if let Some(data) = read_zip_file(zip, &name) {
            if let Ok(xml_str) = std::str::from_utf8(&data) {
                if let Ok(doc) = roxmltree::Document::parse(xml_str) {
                    for node in doc.descendants() {
                        if node.tag_name().name() == "Relationship" {
                            let rid = node.attribute("Id").unwrap_or("");
                            let target = node.attribute("Target").unwrap_or("");
                            if !rid.is_empty() && !target.is_empty() {
                                rels.insert(rid.to_string(), target.to_string());
                            }
                        }
                    }
                }
            }
        }
    }
    rels
}

/// Parse master shapes from all master XML files.
pub fn parse_master_shapes(
    zip: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> HashMap<String, HashMap<String, Shape>> {
    let mut masters: HashMap<String, HashMap<String, Shape>> = HashMap::new();

    // Parse masters.xml to map Master ID -> rel ID -> file
    let mut master_id_to_file: HashMap<String, String> = HashMap::new();

    if let Some(data) = read_zip_file(zip, "visio/masters/masters.xml") {
        if let Ok(xml_str) = std::str::from_utf8(&data) {
            if let Ok(doc) = roxmltree::Document::parse(xml_str) {
                // Parse rels
                let mut rid_to_file: HashMap<String, String> = HashMap::new();
                if let Some(rels_data) = read_zip_file(zip, "visio/masters/_rels/masters.xml.rels")
                {
                    if let Ok(rels_str) = std::str::from_utf8(&rels_data) {
                        if let Ok(rels_doc) = roxmltree::Document::parse(rels_str) {
                            for node in rels_doc.descendants() {
                                if node.tag_name().name() == "Relationship" {
                                    let rid = node.attribute("Id").unwrap_or("");
                                    let target = node.attribute("Target").unwrap_or("");
                                    let fname = std::path::Path::new(target)
                                        .file_stem()
                                        .and_then(|s| s.to_str())
                                        .unwrap_or("")
                                        .replace("master", "");
                                    if !rid.is_empty() {
                                        rid_to_file.insert(rid.to_string(), fname);
                                    }
                                }
                            }
                        }
                    }
                }

                for node in doc.descendants() {
                    if is_visio_tag(&node, "Master") {
                        let mid = node.attribute("ID").unwrap_or("").to_string();
                        if mid.is_empty() {
                            continue;
                        }

                        // Find Rel element
                        let mut mapped = false;
                        for child in node.children() {
                            if child.tag_name().name() == "Rel" {
                                // Try various attribute patterns for r:id
                                let rid = child
                                    .attribute((RNS, "id"))
                                    .or_else(|| child.attribute("id"))
                                    .unwrap_or("");
                                if let Some(fnum) = rid_to_file.get(rid) {
                                    master_id_to_file.insert(mid.clone(), fnum.clone());
                                    mapped = true;
                                }
                                break;
                            }
                        }
                        if !mapped {
                            master_id_to_file.insert(mid.clone(), mid.clone());
                        }
                    }
                }
            }
        }
    }

    // Parse all master files
    let master_file_names: Vec<String> = (0..zip.len())
        .filter_map(|i| {
            let f = zip.by_index(i).ok()?;
            let n = f.name().to_string();
            if n.starts_with("visio/masters/master")
                && n.ends_with(".xml")
                && !n.contains("masters.xml")
            {
                Some(n)
            } else {
                None
            }
        })
        .collect();

    let mut file_to_shapes: HashMap<String, HashMap<String, Shape>> = HashMap::new();

    for name in master_file_names {
        let master_num = std::path::Path::new(&name)
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .replace("master", "");

        if let Some(data) = read_zip_file(zip, &name) {
            if let Ok(xml_str) = std::str::from_utf8(&data) {
                if let Ok(doc) = roxmltree::Document::parse(xml_str) {
                    let mut shapes_data = HashMap::new();
                    for node in doc.descendants() {
                        if is_visio_tag(&node, "Shape") {
                            // Only top-level shapes (parent is not another Shape)
                            let _parent_is_shape = node
                                .parent()
                                .map(|p| is_visio_tag(&p, "Shape") || is_visio_tag(&p, "Shapes"))
                                .unwrap_or(false);
                            let parent_is_shapes_in_shape = node
                                .parent()
                                .and_then(|p| p.parent())
                                .map(|gp| is_visio_tag(&gp, "Shape"))
                                .unwrap_or(false);
                            if parent_is_shapes_in_shape {
                                continue; // Skip sub-shapes, they're parsed recursively
                            }
                            let sd = parse_single_shape(&node);
                            shapes_data.insert(sd.id.clone(), sd);
                        }
                    }
                    if !shapes_data.is_empty() {
                        file_to_shapes.insert(master_num, shapes_data);
                    }
                }
            }
        }
    }

    // Re-key by Master ID
    for (mid, fnum) in &master_id_to_file {
        if let Some(shapes) = file_to_shapes.get(fnum) {
            masters.insert(mid.clone(), shapes.clone());
        }
    }

    // Add unmapped files
    let mapped_files: std::collections::HashSet<&String> = master_id_to_file.values().collect();
    for (fnum, shapes) in &file_to_shapes {
        if !mapped_files.contains(fnum) {
            masters.insert(fnum.clone(), shapes.clone());
        }
    }

    masters
}

/// Parse shapes from a page XML document.
pub fn parse_page_shapes(doc: &roxmltree::Document) -> Vec<Shape> {
    let mut shapes = Vec::new();
    for node in doc.root().children() {
        // Find the root element
        for child in node.children() {
            if is_visio_tag(&child, "Shapes") {
                for shape_node in child.children() {
                    if is_visio_tag(&shape_node, "Shape") {
                        shapes.push(parse_single_shape(&shape_node));
                    }
                }
            }
        }
    }
    shapes
}

/// Parse Connect elements from a page XML.
pub fn parse_connects(doc: &roxmltree::Document) -> Vec<Connect> {
    let mut connects = Vec::new();
    for node in doc.descendants() {
        if is_visio_tag(&node, "Connect") {
            connects.push(Connect {
                from_sheet: node.attribute("FromSheet").unwrap_or("").to_string(),
                from_cell: node.attribute("FromCell").unwrap_or("").to_string(),
                to_sheet: node.attribute("ToSheet").unwrap_or("").to_string(),
                to_cell: node.attribute("ToCell").unwrap_or("").to_string(),
            });
        }
    }
    connects
}

/// Parse layer definitions from a page.
pub fn parse_layers(doc: &roxmltree::Document) -> HashMap<String, LayerDef> {
    let mut layers = HashMap::new();
    for node in doc.descendants() {
        if is_visio_tag(&node, "PageSheet") {
            for section in node.children() {
                if is_visio_tag(&section, "Section") && section.attribute("N") == Some("Layer") {
                    for row in section.children() {
                        if is_visio_tag(&row, "Row") {
                            let ix = row.attribute("IX").unwrap_or("").to_string();
                            let visible = cell_val(&row, "Visible") != "0";
                            let name = cell_val(&row, "Name");
                            let name = if name.is_empty() {
                                format!("Layer {}", ix)
                            } else {
                                name
                            };
                            layers.insert(ix, LayerDef { name, visible });
                        }
                    }
                }
            }
        }
    }
    layers
}

/// Parse a single <Shape> element into a Shape struct.
pub fn parse_single_shape(node: &roxmltree::Node) -> Shape {
    let mut shape = Shape::default();
    shape.id = node.attribute("ID").unwrap_or("").to_string();
    shape.name = node.attribute("Name").unwrap_or("").to_string();
    shape.name_u = node.attribute("NameU").unwrap_or("").to_string();
    shape.shape_type = node.attribute("Type").unwrap_or("Shape").to_string();
    shape.master = node.attribute("Master").unwrap_or("").to_string();
    shape.master_shape = node.attribute("MasterShape").unwrap_or("").to_string();
    shape.line_style = node.attribute("LineStyle").unwrap_or("").to_string();
    shape.fill_style = node.attribute("FillStyle").unwrap_or("").to_string();
    shape.text_style = node.attribute("TextStyle").unwrap_or("").to_string();

    // Parse top-level cells
    for child in node.children() {
        if is_visio_tag(&child, "Cell") {
            let n = child.attribute("N").unwrap_or("");
            let v = child.attribute("V").unwrap_or("");
            let f = child.attribute("F").unwrap_or("");
            shape.cells.insert(n.to_string(), CellValue::new(v, f));
        }
    }

    // Parse sections
    for child in node.children() {
        if is_visio_tag(&child, "Section") {
            let sec_name = child.attribute("N").unwrap_or("");
            match sec_name {
                "Geometry" => {
                    let geo = parse_geometry_section(&child);
                    shape.geometry.push(geo);
                }
                "Character" => {
                    for row in child.children() {
                        if is_visio_tag(&row, "Row") {
                            let ix = row.attribute("IX").unwrap_or("0").to_string();
                            let mut fmt = CharFormat::default();
                            for cell in row.children() {
                                if is_visio_tag(&cell, "Cell") {
                                    let n = cell.attribute("N").unwrap_or("");
                                    let v = cell.attribute("V").unwrap_or("").to_string();
                                    match n {
                                        "Size" => fmt.size = v,
                                        "Color" => fmt.color = v,
                                        "Style" => fmt.style = v,
                                        "Font" => fmt.font = v,
                                        _ => {}
                                    }
                                }
                            }
                            shape.char_formats.insert(ix, fmt);
                        }
                    }
                }
                "Paragraph" => {
                    for row in child.children() {
                        if is_visio_tag(&row, "Row") {
                            let ix = row.attribute("IX").unwrap_or("0").to_string();
                            let mut fmt = ParaFormat::default();
                            for cell in row.children() {
                                if is_visio_tag(&cell, "Cell") {
                                    let n = cell.attribute("N").unwrap_or("");
                                    let v = cell.attribute("V").unwrap_or("").to_string();
                                    match n {
                                        "HorzAlign" => fmt.horiz_align = v,
                                        "IndFirst" => fmt.indent_first = v,
                                        "IndLeft" => fmt.indent_left = v,
                                        "IndRight" => fmt.indent_right = v,
                                        "Bullet" => fmt.bullet = v,
                                        "BulletStr" => fmt.bullet_str = v,
                                        "SpLine" => fmt.sp_line = v,
                                        "SpBefore" => fmt.sp_before = v,
                                        "SpAfter" => fmt.sp_after = v,
                                        _ => {}
                                    }
                                }
                            }
                            shape.para_formats.insert(ix, fmt);
                        }
                    }
                }
                "Controls" => {
                    for row in child.children() {
                        if is_visio_tag(&row, "Row") {
                            let row_ix = format!("Row_{}", row.attribute("IX").unwrap_or("0"));
                            let mut ctrl = HashMap::new();
                            for cell in row.children() {
                                if is_visio_tag(&cell, "Cell") {
                                    ctrl.insert(
                                        cell.attribute("N").unwrap_or("").to_string(),
                                        cell.attribute("V").unwrap_or("").to_string(),
                                    );
                                }
                            }
                            shape.controls.insert(row_ix, ctrl);
                        }
                    }
                }
                "Connection" => {
                    for row in child.children() {
                        if is_visio_tag(&row, "Row") {
                            let ix = row.attribute("IX").unwrap_or("0").to_string();
                            let mut conn = HashMap::new();
                            for cell in row.children() {
                                if is_visio_tag(&cell, "Cell") {
                                    conn.insert(
                                        cell.attribute("N").unwrap_or("").to_string(),
                                        CellValue::new(
                                            cell.attribute("V").unwrap_or(""),
                                            cell.attribute("F").unwrap_or(""),
                                        ),
                                    );
                                }
                            }
                            shape.connections.insert(ix, conn);
                        }
                    }
                }
                "User" => {
                    for row in child.children() {
                        if is_visio_tag(&row, "Row") {
                            let row_name = row.attribute("N").unwrap_or("").to_string();
                            let mut user_vals = HashMap::new();
                            for cell in row.children() {
                                if is_visio_tag(&cell, "Cell") {
                                    user_vals.insert(
                                        cell.attribute("N").unwrap_or("").to_string(),
                                        cell.attribute("V").unwrap_or("").to_string(),
                                    );
                                }
                            }
                            shape.user.insert(row_name, user_vals);
                        }
                    }
                }
                "FillGradientDef" => {
                    let mut stops = Vec::new();
                    for row in child.children() {
                        if is_visio_tag(&row, "Row") {
                            let pos_str = cell_val(&row, "GradientStopPosition");
                            let color = cell_val(&row, "GradientStopColor");
                            let pos: f64 = pos_str.parse().unwrap_or(0.0) * 100.0;
                            if !color.is_empty() {
                                stops.push(GradientStop {
                                    position: pos,
                                    color,
                                });
                            }
                        }
                    }
                    if !stops.is_empty() {
                        shape.gradient_stops.push(stops);
                    }
                }
                "Hyperlink" => {
                    for row in child.children() {
                        if is_visio_tag(&row, "Row") {
                            let mut link = Hyperlink::default();
                            for cell in row.children() {
                                if is_visio_tag(&cell, "Cell") {
                                    let n = cell.attribute("N").unwrap_or("");
                                    let v = cell.attribute("V").unwrap_or("").to_string();
                                    match n {
                                        "Description" => link.description = v,
                                        "Address" => link.address = v,
                                        "SubAddress" => link.sub_address = v,
                                        "Frame" => link.frame = v,
                                        _ => {}
                                    }
                                }
                            }
                            shape.hyperlinks.push(link);
                        }
                    }
                }
                _ => {}
            }
        }
    }

    // Parse Text element
    if let Some(text_node) = find_child(node, "Text") {
        shape.has_text_elem = true;
        let text = collect_text(&text_node);
        shape.text = text.trim().to_string();
        shape.text_parts = parse_text_parts(&text_node);
    }

    // Parse sub-shapes (for groups)
    if let Some(shapes_container) = find_child(node, "Shapes") {
        for sub_node in shapes_container.children() {
            if is_visio_tag(&sub_node, "Shape") {
                shape.sub_shapes.push(parse_single_shape(&sub_node));
            }
        }
    }

    // Parse ForeignData
    if let Some(fd_node) = find_child(node, "ForeignData") {
        let mut fdi = ForeignDataInfo::default();
        fdi.foreign_type = fd_node.attribute("ForeignType").unwrap_or("").to_string();
        fdi.compression = fd_node
            .attribute("CompressionType")
            .unwrap_or("")
            .to_string();

        // Look for Rel element
        let mut found_rel = false;
        for child in fd_node.children() {
            if child.tag_name().name() == "Rel" {
                let rid = child
                    .attribute((RNS, "id"))
                    .or_else(|| child.attribute("id"))
                    .unwrap_or("");
                if !rid.is_empty() {
                    fdi.rel_id = Some(rid.to_string());
                    found_rel = true;
                }
                break;
            }
        }
        if !found_rel {
            if let Some(text) = fd_node.text() {
                let trimmed = text.trim();
                if !trimmed.is_empty() {
                    fdi.data = Some(trimmed.to_string());
                }
            }
        }
        shape.foreign_data = Some(fdi);
    }

    shape
}

fn parse_geometry_section(section: &roxmltree::Node) -> GeomSection {
    let mut geo = GeomSection::default();
    geo.ix = section.attribute("IX").unwrap_or("0").to_string();

    // Section-level cells
    for child in section.children() {
        if is_visio_tag(&child, "Cell") {
            match child.attribute("N") {
                Some("NoFill") if child.attribute("V") == Some("1") => geo.no_fill = true,
                Some("NoLine") if child.attribute("V") == Some("1") => geo.no_line = true,
                Some("NoShow") if child.attribute("V") == Some("1") => geo.no_show = true,
                _ => {}
            }
        }
    }

    for child in section.children() {
        if is_visio_tag(&child, "Row") {
            let row_type = child.attribute("T").unwrap_or("").to_string();
            let row_ix = child.attribute("IX").unwrap_or("").to_string();
            let mut cells = HashMap::new();
            for cell in child.children() {
                if is_visio_tag(&cell, "Cell") {
                    let n = cell.attribute("N").unwrap_or("");
                    let v = cell.attribute("V").unwrap_or("");
                    let f = cell.attribute("F").unwrap_or("");
                    cells.insert(n.to_string(), CellValue::new(v, f));
                }
            }
            geo.rows.push(GeomRow {
                row_type,
                ix: row_ix,
                cells,
            });
        }
    }

    geo
}

fn collect_text(node: &roxmltree::Node) -> String {
    let mut text = String::new();
    if let Some(t) = node.text() {
        text.push_str(t);
    }
    for child in node.children() {
        if child.is_element() {
            if child.tag_name().name() == "fld" {
                text.push_str(&collect_text(&child));
            }
        }
        if let Some(tail) = child.tail() {
            text.push_str(tail);
        }
    }
    text
}

fn parse_text_parts(text_node: &roxmltree::Node) -> Vec<TextPart> {
    let mut parts = Vec::new();
    let mut current_cp = "0".to_string();
    let mut current_pp = "0".to_string();

    if let Some(text) = text_node.text() {
        if !text.is_empty() {
            parts.push(TextPart {
                text: text.to_string(),
                cp: current_cp.clone(),
                pp: current_pp.clone(),
            });
        }
    }

    for child in text_node.children() {
        if child.is_element() {
            match child.tag_name().name() {
                "cp" => {
                    current_cp = child.attribute("IX").unwrap_or("0").to_string();
                }
                "pp" => {
                    current_pp = child.attribute("IX").unwrap_or("0").to_string();
                }
                "fld" => {
                    let field_text = collect_text(&child);
                    let trimmed = field_text.trim();
                    if !trimmed.is_empty() {
                        parts.push(TextPart {
                            text: trimmed.to_string(),
                            cp: current_cp.clone(),
                            pp: current_pp.clone(),
                        });
                    }
                }
                _ => {}
            }
        }
        if let Some(tail) = child.tail() {
            if !tail.is_empty() {
                parts.push(TextPart {
                    text: tail.to_string(),
                    cp: current_cp.clone(),
                    pp: current_pp.clone(),
                });
            }
        }
    }

    parts
}
