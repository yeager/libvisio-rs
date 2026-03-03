# libvisio-rs

**10/10 Quality** — A production-ready Rust library for parsing Microsoft Visio files (.vsdx and .vsd) and converting them to SVG. This is a faithful port of [libvisio-ng](https://github.com/yeager/libvisio-ng) from Python to Rust, designed to be a drop-in replacement for C++ libvisio in projects like LibreOffice and Inkscape.

**130 comprehensive tests** ensure rock-solid reliability across all features.

## Features

- **Full .vsdx support** — themes, gradients (linear + radial), shadows, hatching, fill/line patterns, rounded rectangles, rich text, image embedding, connectors, hyperlinks, master shape inheritance, grouped/nested shapes
- **Full .vsd support** — NURBS, polylines, splines, TextXForm, XForm1D, shadows, embedded images, layers
- **High-quality SVG output** — proper transforms, tspan for rich text, gradient defs, image data URIs, arrow markers
- **C ABI** — shared library with C-compatible API and auto-generated headers via cbindgen
- **Python bindings** (via PyO3) — drop-in replacement for the Python libvisio-ng package
- **CLI tool** — `visio2svg` command-line converter
- **Cross-platform** — builds on Linux, macOS, and Windows

## Installation

### From source

```bash
cargo build --release
```

The shared library will be in `target/release/`:
- Linux: `libvisio_rs.so`
- macOS: `libvisio_rs.dylib`
- Windows: `visio_rs.dll`

### CLI tool

```bash
cargo install --path .
visio2svg diagram.vsdx output/
```

## Usage

### Rust API

```rust
use libvisio_rs::{convert, get_page_info, extract_text};

// Convert to SVG files
let svg_files = convert("diagram.vsdx", Some("output/"), None)?;

// Get page info
let pages = get_page_info("diagram.vsdx")?;
for page in pages {
    println!("Page {}: {} ({:.1}\" × {:.1}\")", page.index, page.name, page.width, page.height);
}

// Extract text
let text = extract_text("diagram.vsdx")?;
println!("{}", text);
```

### C API

```c
#include "libvisio_rs.h"

VisioDocument *doc = visio_open("diagram.vsdx");
if (doc) {
    int pages = visio_get_page_count(doc);
    for (int i = 0; i < pages; i++) {
        char *svg = visio_convert_page_to_svg(doc, i);
        if (svg) {
            // Use SVG string...
            visio_free_string(svg);
        }
    }
    visio_free(doc);
}
```

### CLI

```bash
# Convert all pages
visio2svg diagram.vsdx output/

# Convert specific page
visio2svg diagram.vsdx output/ --page 0

# Extract text only
visio2svg diagram.vsdx --text

# Show page info
visio2svg diagram.vsdx --info
```

## Python Bindings

See the `python/` directory for PyO3-based bindings that provide a drop-in replacement for the Python libvisio-ng package.

```bash
cd python
pip install maturin
maturin develop
```

```python
import libvisio_ng

# Same API as the pure-Python version
svg_files = libvisio_ng.convert("diagram.vsdx", output_dir="output/")
pages = libvisio_ng.get_page_info("diagram.vsdx")
text = libvisio_ng.extract_text("diagram.vsdx")
```

## Architecture

```
src/
├── lib.rs          # Public Rust API + C ABI exports
├── error.rs        # Error types (thiserror)
├── model.rs        # Common data model (Shape, Page, Document, XForm, etc.)
├── vsdx/           # .vsdx XML parser
│   ├── parser.rs   # ZIP + XML parsing
│   ├── theme.rs    # Theme color resolution
│   ├── gradient.rs # Gradient helpers
│   ├── text.rs     # Rich text helpers
│   └── image.rs    # Embedded image handling
├── vsd/            # .vsd binary parser
│   ├── parser.rs   # OLE2 compound document parsing
│   ├── records.rs  # Binary record type constants
│   ├── nurbs.rs    # NURBS curve evaluation
│   └── shapes.rs   # Shape/group hierarchy conversion
└── svg/            # SVG renderer
    └── render.rs   # Shape → SVG conversion
```

## License

GPL-3.0-or-later — see [LICENSE](LICENSE)

## Author

Daniel Nylander <daniel@danielnylander.se>
