//! Common data model for Visio documents.
//!
//! Defines Shape, Page, Document, XForm, and other shared types used by both
//! the .vsdx XML parser and the .vsd binary parser.

use std::collections::HashMap;

/// Transform data for a shape (position, size, rotation, flips).
#[derive(Debug, Clone, Default)]
pub struct XForm {
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

/// Text block transform — positions text independently of shape.
#[derive(Debug, Clone, Default)]
pub struct TextXForm {
    pub txt_pin_x: f64,
    pub txt_pin_y: f64,
    pub txt_width: f64,
    pub txt_height: f64,
    pub txt_loc_pin_x: f64,
    pub txt_loc_pin_y: f64,
    pub txt_angle: f64,
}

/// 1D connector endpoints.
#[derive(Debug, Clone, Default)]
pub struct XForm1D {
    pub begin_x: f64,
    pub begin_y: f64,
    pub end_x: f64,
    pub end_y: f64,
}

/// A cell value with optional formula.
#[derive(Debug, Clone, Default)]
pub struct CellValue {
    pub v: String,
    pub f: String,
}

impl CellValue {
    pub fn new(v: impl Into<String>, f: impl Into<String>) -> Self {
        Self {
            v: v.into(),
            f: f.into(),
        }
    }
    pub fn val(v: impl Into<String>) -> Self {
        Self {
            v: v.into(),
            f: String::new(),
        }
    }
    pub fn as_f64(&self) -> f64 {
        self.v.parse().unwrap_or(0.0)
    }
    pub fn as_f64_or(&self, default: f64) -> f64 {
        self.v.parse().unwrap_or(default)
    }
    pub fn is_empty(&self) -> bool {
        self.v.is_empty() && self.f.is_empty()
    }
}

/// A single row in a geometry section.
#[derive(Debug, Clone)]
pub struct GeomRow {
    pub row_type: String,
    pub ix: String,
    pub cells: HashMap<String, CellValue>,
}

impl GeomRow {
    pub fn new(row_type: &str) -> Self {
        Self {
            row_type: row_type.to_string(),
            ix: String::new(),
            cells: HashMap::new(),
        }
    }
    pub fn cell_f64(&self, name: &str) -> f64 {
        self.cells.get(name).map(|c| c.as_f64()).unwrap_or(0.0)
    }
}

/// A geometry section (may contain multiple rows).
#[derive(Debug, Clone, Default)]
pub struct GeomSection {
    pub no_fill: bool,
    pub no_line: bool,
    pub no_show: bool,
    pub ix: String,
    pub rows: Vec<GeomRow>,
}

/// Character formatting for a text run.
#[derive(Debug, Clone)]
pub struct CharFormat {
    pub size: String,
    pub color: String,
    pub style: String,
    pub font: String,
}

impl Default for CharFormat {
    fn default() -> Self {
        Self {
            size: "0.1111".to_string(),
            color: String::new(),
            style: "0".to_string(),
            font: String::new(),
        }
    }
}

/// Paragraph formatting.
#[derive(Debug, Clone, Default)]
pub struct ParaFormat {
    pub horiz_align: String,
    pub indent_first: String,
    pub indent_left: String,
    pub indent_right: String,
    pub bullet: String,
    pub bullet_str: String,
    pub sp_line: String,
    pub sp_before: String,
    pub sp_after: String,
}

/// A part of text with formatting references.
#[derive(Debug, Clone)]
pub struct TextPart {
    pub text: String,
    pub cp: String,
    pub pp: String,
}

/// Foreign data info (embedded images).
#[derive(Debug, Clone, Default)]
pub struct ForeignDataInfo {
    pub foreign_type: String,
    pub compression: String,
    pub data: Option<String>,
    pub rel_id: Option<String>,
}

/// Hyperlink data.
#[derive(Debug, Clone, Default)]
pub struct Hyperlink {
    pub description: String,
    pub address: String,
    pub sub_address: String,
    pub frame: String,
}

/// Gradient stop.
#[derive(Debug, Clone)]
pub struct GradientStop {
    pub position: f64,
    pub color: String,
}

/// Gradient definition for SVG output.
#[derive(Debug, Clone)]
pub struct GradientDef {
    pub id: String,
    pub start: String,
    pub end: String,
    pub mid: Option<String>,
    pub dir: f64,
    pub radial: bool,
    pub stops: Vec<GradientStop>,
    pub cx: f64,
    pub cy: f64,
    pub fx: f64,
    pub fy: f64,
    pub r: f64,
}

impl Default for GradientDef {
    fn default() -> Self {
        Self {
            id: String::new(),
            start: "#FFFFFF".to_string(),
            end: "#000000".to_string(),
            mid: None,
            dir: 0.0,
            radial: false,
            stops: Vec::new(),
            cx: 50.0,
            cy: 50.0,
            fx: 50.0,
            fy: 50.0,
            r: 50.0,
        }
    }
}

/// Fill pattern definition for hatching.
#[derive(Debug, Clone)]
pub struct FillPatternDef {
    pub id: String,
    pub fg: String,
    pub bg: String,
    pub pattern_type: i32,
}

/// A parsed Visio shape.
#[derive(Debug, Clone)]
pub struct Shape {
    pub id: String,
    pub name: String,
    pub name_u: String,
    pub shape_type: String,
    pub master: String,
    pub master_shape: String,
    pub cells: HashMap<String, CellValue>,
    pub geometry: Vec<GeomSection>,
    pub text: String,
    pub text_parts: Vec<TextPart>,
    pub char_formats: HashMap<String, CharFormat>,
    pub para_formats: HashMap<String, ParaFormat>,
    pub sub_shapes: Vec<Shape>,
    pub controls: HashMap<String, HashMap<String, String>>,
    pub connections: HashMap<String, HashMap<String, CellValue>>,
    pub user: HashMap<String, HashMap<String, String>>,
    pub foreign_data: Option<ForeignDataInfo>,
    pub hyperlinks: Vec<Hyperlink>,
    pub line_style: String,
    pub fill_style: String,
    pub text_style: String,
    pub gradient_stops: Vec<Vec<GradientStop>>,
    pub has_text_elem: bool,
    pub master_w: f64,
    pub master_h: f64,
    pub has_own_geometry: bool,
    pub theme_text_color: Option<String>,
}

impl Default for Shape {
    fn default() -> Self {
        Self {
            id: String::new(),
            name: String::new(),
            name_u: String::new(),
            shape_type: "Shape".to_string(),
            master: String::new(),
            master_shape: String::new(),
            cells: HashMap::new(),
            geometry: Vec::new(),
            text: String::new(),
            text_parts: Vec::new(),
            char_formats: HashMap::new(),
            para_formats: HashMap::new(),
            sub_shapes: Vec::new(),
            controls: HashMap::new(),
            connections: HashMap::new(),
            user: HashMap::new(),
            foreign_data: None,
            hyperlinks: Vec::new(),
            line_style: String::new(),
            fill_style: String::new(),
            text_style: String::new(),
            gradient_stops: Vec::new(),
            has_text_elem: false,
            master_w: 0.0,
            master_h: 0.0,
            has_own_geometry: false,
            theme_text_color: None,
        }
    }
}

impl Shape {
    pub fn cell_val(&self, name: &str) -> &str {
        self.cells.get(name).map(|c| c.v.as_str()).unwrap_or("")
    }

    pub fn cell_f64(&self, name: &str) -> f64 {
        self.cells.get(name).map(|c| c.as_f64()).unwrap_or(0.0)
    }

    pub fn cell_f64_or(&self, name: &str, default: f64) -> f64 {
        self.cells
            .get(name)
            .map(|c| c.as_f64_or(default))
            .unwrap_or(default)
    }
}

/// A connection between shapes on a page.
#[derive(Debug, Clone)]
pub struct Connect {
    pub from_sheet: String,
    pub from_cell: String,
    pub to_sheet: String,
    pub to_cell: String,
}

/// Layer definition on a page.
#[derive(Debug, Clone)]
pub struct LayerDef {
    pub name: String,
    pub visible: bool,
}

/// A page in a Visio document.
#[derive(Debug, Clone)]
pub struct Page {
    pub name: String,
    pub index: usize,
    pub width: f64,
    pub height: f64,
    pub shapes: Vec<Shape>,
    pub connects: Vec<Connect>,
    pub layers: HashMap<String, LayerDef>,
    pub background: bool,
}

impl Default for Page {
    fn default() -> Self {
        Self {
            name: String::new(),
            index: 0,
            width: 8.5,
            height: 11.0,
            shapes: Vec::new(),
            connects: Vec::new(),
            layers: HashMap::new(),
            background: false,
        }
    }
}

/// Stylesheet data from document.xml.
#[derive(Debug, Clone, Default)]
pub struct StyleSheet {
    pub cells: HashMap<String, CellValue>,
    pub line_style: String,
    pub fill_style: String,
    pub text_style: String,
}

/// A parsed Visio document.
#[derive(Debug, Clone, Default)]
pub struct Document {
    pub pages: Vec<Page>,
    pub masters: HashMap<String, HashMap<String, Shape>>,
    pub theme_colors: HashMap<String, String>,
    pub media: HashMap<String, Vec<u8>>,
    pub stylesheets: HashMap<String, StyleSheet>,
    pub background_map: HashMap<usize, usize>,
}

/// Page information returned by get_page_info.
#[derive(Debug, Clone)]
pub struct PageInfo {
    pub name: String,
    pub index: usize,
    pub width: f64,
    pub height: f64,
}
