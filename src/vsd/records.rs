//! Binary record type constants for .vsd files.
//! Based on libvisio VSDDocumentStructure.h from LibreOffice.

pub const VSD_FOREIGN_DATA: u32 = 0x0C;
pub const VSD_OLE_LIST: u32 = 0x0D;
pub const VSD_TEXT: u32 = 0x0E;
pub const VSD_PAGE: u32 = 0x15;
pub const VSD_COLORS: u32 = 0x16;
pub const VSD_FONT_IX: u32 = 0x19;
pub const VSD_STENCILS: u32 = 0x1D;
pub const VSD_STENCIL_PAGE: u32 = 0x1E;
pub const VSD_OLE_DATA: u32 = 0x1F;
pub const VSD_PAGES: u32 = 0x27;
pub const VSD_NAME_LIST2: u32 = 0x32;
pub const VSD_NAME2: u32 = 0x33;
pub const VSD_NAMEIDX123: u32 = 0x34;
pub const VSD_PAGE_SHEET: u32 = 0x46;
pub const VSD_SHAPE_GROUP: u32 = 0x47;
pub const VSD_SHAPE_SHAPE: u32 = 0x48;
pub const VSD_SHAPE_FOREIGN: u32 = 0x4E;
pub const VSD_SHAPE_LIST: u32 = 0x65;
pub const VSD_CHAR_LIST: u32 = 0x69;
pub const VSD_PARA_LIST: u32 = 0x6A;
pub const VSD_GEOM_LIST: u32 = 0x6C;
pub const VSD_SHAPE_ID: u32 = 0x83;
pub const VSD_LINE: u32 = 0x85;
pub const VSD_FILL_AND_SHADOW: u32 = 0x86;
pub const VSD_TEXT_BLOCK: u32 = 0x87;
pub const VSD_GEOMETRY: u32 = 0x89;
pub const VSD_MOVE_TO: u32 = 0x8A;
pub const VSD_LINE_TO: u32 = 0x8B;
pub const VSD_ARC_TO: u32 = 0x8C;
pub const VSD_ELLIPSE: u32 = 0x8F;
pub const VSD_ELLIPTICAL_ARC_TO: u32 = 0x90;
pub const VSD_PAGE_PROPS: u32 = 0x92;
pub const VSD_CHAR_IX: u32 = 0x94;
pub const VSD_PARA_IX: u32 = 0x95;
pub const VSD_FOREIGN_DATA_TYPE: u32 = 0x98;
pub const VSD_CONNECTION_POINTS: u32 = 0x99;
pub const VSD_XFORM_DATA: u32 = 0x9B;
pub const VSD_TEXT_XFORM: u32 = 0x9C;
pub const VSD_XFORM_1D: u32 = 0x9D;
pub const VSD_SPLINE_START: u32 = 0xA5;
pub const VSD_SPLINE_KNOT: u32 = 0xA6;
pub const VSD_LAYER_MEMBERSHIP: u32 = 0xA7;
pub const VSD_INFINITE_LINE: u32 = 0x8D;
pub const VSD_POLYLINE_TO: u32 = 0xC1;
pub const VSD_NURBS_TO: u32 = 0xC3;
pub const VSD_NAMEIDX: u32 = 0xC9;
pub const VSD_FONTFACE: u32 = 0xD7;
pub const VSD_FONTFACES: u32 = 0xD8;
pub const VSD_NAME: u32 = 0x2D;

/// Chunk header for .vsd binary records.
#[derive(Debug, Clone, Default)]
pub struct ChunkHeader {
    pub chunk_type: u32,
    pub record_id: u32,
    pub list_flag: u32,
    pub data_length: u32,
    pub level: u16,
    pub unknown: u8,
    pub trailer: u32,
}

/// Pointer in the trailer/pointer tree.
#[derive(Debug, Clone, Default)]
pub struct Pointer {
    pub ptr_type: u32,
    pub offset: u32,
    pub length: u32,
    pub fmt: u16,
}

/// Types that get trailer bytes.
pub static TRAILER_TYPES: &[u32] = &[
    0x64, 0x65, 0x66, 0x69, 0x6A, 0x6B, 0x6F, 0x71, 0x92, 0xA9, 0xB4, 0xB6, 0xB9, 0xC7,
];

/// List types that get trailer bytes.
pub static LIST_TRAILER_TYPES: &[u32] = &[0x71, 0x70, 0x6B, 0x6A, 0x69, 0x66, 0x65, 0x2C];

/// Types that never get trailers.
pub static NO_TRAILER_TYPES: &[u32] = &[0x1F, 0xC9, 0x2D, 0xD1];
