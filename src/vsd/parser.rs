//! .vsd (OLE2/Compound Binary) parser.
//!
//! Reads OLE2 structured storage and parses the VisioDocument stream
//! containing pointer-based tree of binary records.

use crate::error::{Result, VisioError};
use crate::model::*;
use crate::vsd::records::*;
use crate::vsd::shapes::*;
use std::io::Read;

/// Parse a .vsd file from bytes.
pub fn parse_vsd(data: &[u8]) -> Result<Document> {
    let cursor = std::io::Cursor::new(data);
    let mut comp = cfb::CompoundFile::open(cursor).map_err(|e| VisioError::Cfb(e.to_string()))?;

    let mut stream_data = Vec::new();
    {
        let mut stream = comp
            .open_stream("/VisioDocument")
            .map_err(|e| VisioError::Cfb(format!("Cannot open VisioDocument stream: {}", e)))?;
        stream
            .read_to_end(&mut stream_data)
            .map_err(|e| VisioError::Io(e))?;
    }

    let mut parser = VsdParser::new(&stream_data);
    parser.parse()?;
    Ok(parser.into_document())
}

#[allow(dead_code)]
struct VsdParser<'a> {
    data: &'a [u8],
    pages: Vec<Page>,
    current_page: Option<Page>,
    current_shape: Option<CurrentShape>,
    current_geom: Option<VsdGeomSection>,
    colors: Vec<String>, // reserved for future use,
    fonts: std::collections::HashMap<u32, String>,
    names: std::collections::HashMap<u32, String>,
}

struct CurrentShape {
    shape_id: u32,
    shape_type: String,
    xform: VsdXForm,
    text_xform: Option<VsdTextXForm>,
    xform_1d: Option<VsdXForm1D>,
    text: String,
    geometry: Vec<VsdGeomSection>,
    char_formats: Vec<VsdCharFormat>,
    para_formats: Vec<VsdParaFormat>,
    line_weight: f64,
    line_color: String,
    line_pattern: i32,
    fill_fg: String,
    fill_bg: String,
    fill_pattern: i32,
    shadow_color: String,
    shadow_pattern: i32,
    shadow_offset_x: f64,
    shadow_offset_y: f64,
    foreign_data: Option<VsdForeignData>,
    layer_member: String,
}

impl Default for CurrentShape {
    fn default() -> Self {
        Self {
            shape_id: 0,
            shape_type: "Shape".to_string(),
            xform: VsdXForm::default(),
            text_xform: None,
            xform_1d: None,
            text: String::new(),
            geometry: Vec::new(),
            char_formats: Vec::new(),
            para_formats: Vec::new(),
            line_weight: 0.01,
            line_color: "#000000".to_string(),
            line_pattern: 1,
            fill_fg: String::new(),
            fill_bg: String::new(),
            fill_pattern: 1,
            shadow_color: String::new(),
            shadow_pattern: 0,
            shadow_offset_x: 0.0,
            shadow_offset_y: 0.0,
            foreign_data: None,
            layer_member: String::new(),
        }
    }
}

impl<'a> VsdParser<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self {
            data,
            pages: Vec::new(),
            current_page: None,
            current_shape: None,
            current_geom: None,
            colors: Vec::new(),
            fonts: std::collections::HashMap::new(),
            names: std::collections::HashMap::new(),
        }
    }

    fn parse(&mut self) -> Result<()> {
        if self.data.len() < 0x36 {
            return Err(VisioError::InvalidFile("File too small".to_string()));
        }
        // Try linear chunk scanning from offset 0x36
        self.parse_chunks_linear(0x36);
        self.flush_shape();
        Ok(())
    }

    fn into_document(self) -> Document {
        let mut doc = Document::default();
        doc.pages = self.pages;
        doc
    }

    fn flush_shape(&mut self) {
        if let Some(cs) = self.current_shape.take() {
            let mut geom = cs.geometry;
            if let Some(g) = self.current_geom.take() {
                geom.push(g);
            }
            let shape = vsd_shape_to_model(
                &cs.xform,
                &cs.text,
                &geom,
                cs.line_weight,
                &cs.line_color,
                cs.line_pattern,
                &cs.fill_fg,
                &cs.fill_bg,
                cs.fill_pattern,
                cs.shape_id,
                &cs.shape_type,
                cs.text_xform.as_ref(),
                cs.xform_1d.as_ref(),
                &cs.char_formats,
                &cs.para_formats,
                &cs.shadow_color,
                cs.shadow_pattern,
                cs.shadow_offset_x,
                cs.shadow_offset_y,
                cs.foreign_data.as_ref(),
                &cs.layer_member,
            );
            if let Some(page) = &mut self.current_page {
                page.shapes.push(shape);
            }
        }
        self.current_geom = None;
    }

    fn parse_chunks_linear(&mut self, start: usize) {
        let mut offset = start;
        while offset + 19 < self.data.len() {
            // Skip null bytes
            while offset < self.data.len() && self.data[offset] == 0 {
                offset += 1;
            }
            if offset + 19 > self.data.len() {
                break;
            }

            let hdr = match parse_chunk_header(self.data, offset) {
                Some((h, new_off)) => {
                    offset = new_off;
                    h
                }
                None => break,
            };

            let end_pos = offset + hdr.data_length as usize + hdr.trailer as usize;
            if end_pos > self.data.len() {
                break;
            }

            let chunk_data = &self.data[offset..offset + hdr.data_length as usize];
            self.handle_chunk(&hdr, chunk_data);
            offset = end_pos;
        }
    }

    fn handle_chunk(&mut self, hdr: &ChunkHeader, data: &[u8]) {
        match hdr.chunk_type {
            VSD_SHAPE_GROUP | VSD_SHAPE_SHAPE | VSD_SHAPE_FOREIGN => {
                self.flush_shape();
                let mut cs = CurrentShape::default();
                cs.shape_id = hdr.record_id;
                if hdr.chunk_type == VSD_SHAPE_GROUP {
                    cs.shape_type = "Group".to_string();
                } else if hdr.chunk_type == VSD_SHAPE_FOREIGN {
                    cs.shape_type = "Foreign".to_string();
                }
                self.current_shape = Some(cs);
            }
            VSD_PAGE_PROPS => self.read_page_props(data),
            VSD_XFORM_DATA => self.read_xform(data),
            VSD_TEXT_XFORM => self.read_text_xform(data),
            VSD_XFORM_1D => self.read_xform_1d(data),
            VSD_TEXT => self.read_text(data),
            VSD_GEOMETRY => self.read_geometry(data),
            VSD_MOVE_TO => self.read_geom_row("MoveTo", data),
            VSD_LINE_TO => self.read_geom_row("LineTo", data),
            VSD_ARC_TO => self.read_geom_row("ArcTo", data),
            VSD_ELLIPSE => self.read_geom_row("Ellipse", data),
            VSD_ELLIPTICAL_ARC_TO => self.read_geom_row("EllipticalArcTo", data),
            VSD_SPLINE_START => self.read_geom_row("SplineStart", data),
            VSD_SPLINE_KNOT => self.read_geom_row("SplineKnot", data),
            VSD_INFINITE_LINE => self.read_geom_row("InfiniteLine", data),
            VSD_NURBS_TO => self.read_nurbs_to(data),
            VSD_POLYLINE_TO => self.read_polyline_to(data),
            VSD_LINE => self.read_line_fmt(data),
            VSD_FILL_AND_SHADOW => self.read_fill(data),
            VSD_CHAR_IX => self.read_char_ix(data),
            VSD_PARA_IX => self.read_para_ix(data),
            VSD_LAYER_MEMBERSHIP => self.read_layer_membership(data),
            VSD_FOREIGN_DATA_TYPE => self.read_foreign_data_type(data),
            VSD_FOREIGN_DATA => self.read_foreign_data(data),
            VSD_PAGE => {
                self.flush_shape();
                let page = Page::default();
                self.current_page = Some(page);
            }
            VSD_PAGE_SHEET => {
                if self.current_page.is_none() {
                    self.current_page = Some(Page::default());
                }
            }
            _ => {}
        }

        // After page is complete, push it
        if hdr.chunk_type == VSD_PAGE && self.current_page.is_some() {
            // Page will be pushed when the next page starts or at end
        }
    }

    fn read_page_props(&mut self, data: &[u8]) {
        if self.current_page.is_none() {
            self.current_page = Some(Page::default());
        }
        let page = self.current_page.as_mut().unwrap();
        let mut off = 1usize;
        if let Some(w) = read_double(data, off) {
            page.width = w;
        }
        off += 9;
        if let Some(h) = read_double(data, off) {
            page.height = h;
        }
    }

    fn read_xform(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        let mut off = 1usize;
        if let Some(v) = read_double(data, off) {
            cs.xform.pin_x = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            cs.xform.pin_y = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            cs.xform.width = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            cs.xform.height = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            cs.xform.loc_pin_x = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            cs.xform.loc_pin_y = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            cs.xform.angle = v;
        }
        off += 8;
        if off < data.len() {
            cs.xform.flip_x = data[off] != 0;
            off += 1;
        }
        if off < data.len() {
            cs.xform.flip_y = data[off] != 0;
        }
    }

    fn read_text_xform(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        let mut txf = VsdTextXForm::default();
        let mut off = 1usize;
        if let Some(v) = read_double(data, off) {
            txf.txt_pin_x = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            txf.txt_pin_y = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            txf.txt_width = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            txf.txt_height = v;
        }
        cs.text_xform = Some(txf);
    }

    fn read_xform_1d(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        let mut xf = VsdXForm1D::default();
        let mut off = 1usize;
        if let Some(v) = read_double(data, off) {
            xf.begin_x = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            xf.begin_y = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            xf.end_x = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            xf.end_y = v;
        }
        cs.xform_1d = Some(xf);
    }

    fn read_text(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        if data.len() < 8 {
            return;
        }
        let text_data = &data[8..];
        // Try UTF-16LE first
        if text_data.len() >= 2 {
            let chars: Vec<u16> = text_data
                .chunks(2)
                .map(|c| {
                    if c.len() == 2 {
                        u16::from_le_bytes([c[0], c[1]])
                    } else {
                        0
                    }
                })
                .collect();
            if let Ok(s) = String::from_utf16(&chars) {
                let trimmed = s.trim_end_matches('\0').to_string();
                if !trimmed.is_empty() && !trimmed.chars().all(|c| c == '\u{FFFD}') {
                    cs.text = trimmed;
                    return;
                }
            }
        }
        // Fallback to UTF-8
        cs.text = String::from_utf8_lossy(text_data)
            .trim_end_matches('\0')
            .to_string();
    }

    fn read_geometry(&mut self, data: &[u8]) {
        if self.current_shape.is_none() {
            return;
        }
        // Push previous geometry section
        if let Some(g) = self.current_geom.take() {
            if let Some(cs) = &mut self.current_shape {
                cs.geometry.push(g);
            }
        }
        let mut geom = VsdGeomSection::default();
        if !data.is_empty() {
            let flags = data[0];
            geom.no_fill = flags & 1 != 0;
            geom.no_line = flags & 2 != 0;
            geom.no_show = flags & 4 != 0;
        }
        self.current_geom = Some(geom);
    }

    fn ensure_geom(&mut self) -> bool {
        if self.current_shape.is_none() {
            return false;
        }
        if self.current_geom.is_none() {
            self.current_geom = Some(VsdGeomSection::default());
        }
        true
    }

    fn read_geom_row(&mut self, row_type: &str, data: &[u8]) {
        if !self.ensure_geom() {
            return;
        }
        let mut row = VsdGeomRow::default();
        row.row_type = row_type.to_string();
        let mut off = 1usize;
        if let Some(v) = read_double(data, off) {
            row.x = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            row.y = v;
        }
        off += 9;
        match row_type {
            "ArcTo" | "SplineStart" | "SplineKnot" | "InfiniteLine" => {
                if let Some(v) = read_double(data, off) {
                    row.a = v;
                }
                off += 9;
                if let Some(v) = read_double(data, off) {
                    row.b = v;
                }
                off += 9;
            }
            "Ellipse" | "EllipticalArcTo" => {
                if let Some(v) = read_double(data, off) {
                    row.a = v;
                }
                off += 9;
                if let Some(v) = read_double(data, off) {
                    row.b = v;
                }
                off += 9;
                if let Some(v) = read_double(data, off) {
                    row.c = v;
                }
                off += 9;
                if let Some(v) = read_double(data, off) {
                    row.d = v;
                }
            }
            _ => {}
        }
        if let Some(geom) = &mut self.current_geom {
            geom.rows.push(row);
        }
    }

    fn read_nurbs_to(&mut self, data: &[u8]) {
        if !self.ensure_geom() {
            return;
        }
        let mut row = VsdGeomRow::default();
        row.row_type = "NURBSTo".to_string();
        let mut off = 1usize;
        if let Some(v) = read_double(data, off) {
            row.x = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            row.y = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            row.knot_last = v;
        }
        off += 9;
        if off + 2 <= data.len() {
            row.degree = u16::from_le_bytes([data[off], data[off + 1]]);
            off += 2;
        }
        if off < data.len() {
            row.x_type = data[off];
            off += 1;
        }
        if off < data.len() {
            row.y_type = data[off];
            off += 1;
        }
        // Read control points
        while off + 32 <= data.len() {
            let knot = read_double(data, off).unwrap_or(0.0);
            off += 8;
            let weight = read_double(data, off).unwrap_or(0.0);
            off += 8;
            let px = read_double(data, off).unwrap_or(0.0);
            off += 8;
            let py = read_double(data, off).unwrap_or(0.0);
            off += 8;
            row.points.push((px, py, knot, weight));
        }
        if let Some(geom) = &mut self.current_geom {
            geom.rows.push(row);
        }
    }

    fn read_polyline_to(&mut self, data: &[u8]) {
        if !self.ensure_geom() {
            return;
        }
        let mut row = VsdGeomRow::default();
        row.row_type = "PolylineTo".to_string();
        let mut off = 1usize;
        if let Some(v) = read_double(data, off) {
            row.x = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            row.y = v;
        }
        off += 9;
        if off < data.len() {
            row.x_type = data[off];
            off += 1;
        }
        if off < data.len() {
            row.y_type = data[off];
            off += 1;
        }
        while off + 16 <= data.len() {
            let px = read_double(data, off).unwrap_or(0.0);
            off += 8;
            let py = read_double(data, off).unwrap_or(0.0);
            off += 8;
            row.points.push((px, py, 0.0, 0.0));
        }
        if let Some(geom) = &mut self.current_geom {
            geom.rows.push(row);
        }
    }

    fn read_line_fmt(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        let mut off = 1usize;
        if let Some(v) = read_double(data, off) {
            cs.line_weight = v;
        }
        off += 9;
        if off + 3 <= data.len() {
            let r = data[off];
            let g = data[off + 1];
            let b = data[off + 2];
            cs.line_color = format!("#{:02X}{:02X}{:02X}", r, g, b);
            off += 4; // skip alpha
        }
        if off < data.len() {
            cs.line_pattern = data[off] as i32;
        }
    }

    fn read_fill(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        let mut off = 1usize;
        if off + 3 <= data.len() {
            cs.fill_fg = format!(
                "#{:02X}{:02X}{:02X}",
                data[off],
                data[off + 1],
                data[off + 2]
            );
            off += 4;
        }
        off += 1;
        if off + 3 <= data.len() {
            cs.fill_bg = format!(
                "#{:02X}{:02X}{:02X}",
                data[off],
                data[off + 1],
                data[off + 2]
            );
            off += 4;
        }
        if off < data.len() {
            cs.fill_pattern = data[off] as i32;
            off += 1;
        }
        // Shadow data
        if off + 12 <= data.len() {
            off += 1;
            if off + 3 <= data.len() {
                cs.shadow_color = format!(
                    "#{:02X}{:02X}{:02X}",
                    data[off],
                    data[off + 1],
                    data[off + 2]
                );
                off += 4;
            }
            if off < data.len() {
                cs.shadow_pattern = data[off] as i32;
                off += 1;
            }
            off += 1;
            if let Some(v) = read_double(data, off) {
                cs.shadow_offset_x = v;
            }
            off += 9;
            if let Some(v) = read_double(data, off) {
                cs.shadow_offset_y = v;
            }
        }
    }

    fn read_char_ix(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        if data.len() < 12 {
            return;
        }
        let mut fmt = VsdCharFormat::default();
        let mut off = 0usize;
        fmt.char_count = read_u32(data, off);
        off += 4;
        fmt.font_id = read_u16(data, off);
        off += 3;
        if off + 3 <= data.len() {
            fmt.color_r = data[off];
            fmt.color_g = data[off + 1];
            fmt.color_b = data[off + 2];
            off += 4;
        }
        if off < data.len() {
            let mods = data[off];
            fmt.bold = mods & 1 != 0;
            fmt.italic = mods & 2 != 0;
            fmt.underline = mods & 4 != 0;
            off += 5;
        }
        if let Some(v) = read_double(data, off) {
            fmt.font_size = v;
        }
        cs.char_formats.push(fmt);
    }

    fn read_para_ix(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        if data.len() < 8 {
            return;
        }
        let mut pf = VsdParaFormat::default();
        let mut off = 0usize;
        pf.char_count = read_u32(data, off);
        off += 5;
        if let Some(v) = read_double(data, off) {
            pf.indent_first = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            pf.indent_left = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            pf.indent_right = v;
        }
        off += 9;
        if let Some(v) = read_double(data, off) {
            pf.spacing_line = v;
        }
        off += 9;
        // Skip spacing before/after
        off += 18;
        off += 1;
        if off < data.len() {
            pf.horiz_align = data[off];
            off += 2;
        }
        if off < data.len() {
            pf.bullet = data[off];
        }
        cs.para_formats.push(pf);
    }

    fn read_layer_membership(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        if data.len() < 2 {
            return;
        }
        let chars: Vec<u16> = data
            .chunks(2)
            .map(|c| {
                if c.len() == 2 {
                    u16::from_le_bytes([c[0], c[1]])
                } else {
                    0
                }
            })
            .collect();
        if let Ok(s) = String::from_utf16(&chars) {
            cs.layer_member = s.trim_end_matches('\0').trim().to_string();
        }
    }

    fn read_foreign_data_type(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        if cs.foreign_data.is_none() {
            cs.foreign_data = Some(VsdForeignData::default());
        }
        // Skip image offset/dims, read image type
        let off = 34; // Skip 4 doubles + 2 bytes
        if off + 2 <= data.len() {
            let img_type = u16::from_le_bytes([data[off], data[off + 1]]);
            let fd = cs.foreign_data.as_mut().unwrap();
            let (dt, fmt) = match img_type {
                0 => ("img", "emf"),
                1 => ("img", "wmf"),
                2 => ("img", "bmp"),
                3 => ("ole", "ole"),
                4 => ("img", "jpg"),
                5 => ("img", "png"),
                6 => ("img", "gif"),
                7 => ("img", "tiff"),
                _ => ("img", "png"),
            };
            fd.data_type = dt.to_string();
            fd.img_format = fmt.to_string();
        }
    }

    fn read_foreign_data(&mut self, data: &[u8]) {
        let cs = match &mut self.current_shape {
            Some(s) => s,
            None => return,
        };
        if cs.foreign_data.is_none() {
            cs.foreign_data = Some(VsdForeignData::default());
        }
        cs.foreign_data.as_mut().unwrap().data = data.to_vec();
    }
}

// Helper functions for binary reading

fn read_double(data: &[u8], offset: usize) -> Option<f64> {
    if offset + 8 > data.len() {
        return None;
    }
    Some(f64::from_le_bytes(
        data[offset..offset + 8].try_into().ok()?,
    ))
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    if offset + 4 > data.len() {
        return 0;
    }
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap_or([0; 4]))
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    if offset + 2 > data.len() {
        return 0;
    }
    u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap_or([0; 2]))
}

fn parse_chunk_header(data: &[u8], offset: usize) -> Option<(ChunkHeader, usize)> {
    if offset + 19 > data.len() {
        return None;
    }
    let mut off = offset;
    let mut hdr = ChunkHeader::default();
    hdr.chunk_type = read_u32(data, off);
    off += 4;
    hdr.record_id = read_u32(data, off);
    off += 4;
    hdr.list_flag = read_u32(data, off);
    off += 4;

    if hdr.list_flag != 0 || LIST_TRAILER_TYPES.contains(&hdr.chunk_type) {
        hdr.trailer += 8;
    }

    hdr.data_length = read_u32(data, off);
    off += 4;
    hdr.level = read_u16(data, off);
    off += 2;
    hdr.unknown = if off < data.len() { data[off] } else { 0 };
    off += 1;

    if hdr.list_flag != 0
        || (hdr.level == 2 && hdr.unknown == 0x55)
        || (hdr.level == 2 && hdr.unknown == 0x54 && hdr.chunk_type == 0xAA)
        || (hdr.level == 3 && hdr.unknown != 0x50 && hdr.unknown != 0x54)
    {
        hdr.trailer += 4;
    }

    for &tt in TRAILER_TYPES {
        if hdr.chunk_type == tt && hdr.trailer != 12 && hdr.trailer != 4 {
            hdr.trailer += 4;
            break;
        }
    }

    if NO_TRAILER_TYPES.contains(&hdr.chunk_type) {
        hdr.trailer = 0;
    }

    Some((hdr, off))
}
