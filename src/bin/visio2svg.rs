//! visio2svg — CLI tool for converting Visio files to SVG.

use std::process;

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() < 2 || args.contains(&"--help".to_string()) || args.contains(&"-h".to_string()) {
        eprintln!("visio2svg — Convert Visio files (.vsdx/.vsd) to SVG");
        eprintln!();
        eprintln!("Usage: visio2svg <input> [output_dir] [--page N] [--text] [--info]");
        eprintln!();
        eprintln!("Options:");
        eprintln!("  --page N    Convert only page N (0-based)");
        eprintln!("  --text      Extract text only (no SVG)");
        eprintln!("  --info      Show page info only");
        eprintln!("  --help      Show this help");
        process::exit(1);
    }

    let input = &args[1];
    let mut output_dir = None;
    let mut page = None;
    let mut text_mode = false;
    let mut info_mode = false;

    let mut i = 2;
    while i < args.len() {
        match args[i].as_str() {
            "--page" => {
                i += 1;
                if i < args.len() {
                    page = args[i].parse().ok();
                }
            }
            "--text" => text_mode = true,
            "--info" => info_mode = true,
            _ => {
                if !args[i].starts_with('-') && output_dir.is_none() {
                    output_dir = Some(args[i].clone());
                }
            }
        }
        i += 1;
    }

    if !libvisio_rs::is_supported(input) {
        eprintln!("Error: Unsupported file format");
        process::exit(1);
    }

    if text_mode {
        match libvisio_rs::extract_text(input) {
            Ok(text) => println!("{}", text),
            Err(e) => {
                eprintln!("Error: {}", e);
                process::exit(1);
            }
        }
        return;
    }

    if info_mode {
        match libvisio_rs::get_page_info(input) {
            Ok(pages) => {
                for p in pages {
                    println!(
                        "Page {}: \"{}\" ({:.2}\" × {:.2}\")",
                        p.index, p.name, p.width, p.height
                    );
                }
            }
            Err(e) => {
                eprintln!("Error: {}", e);
                process::exit(1);
            }
        }
        return;
    }

    let out = output_dir.as_deref().unwrap_or(".");
    match libvisio_rs::convert(input, Some(out), page) {
        Ok(files) => {
            for f in &files {
                println!("{}", f);
            }
            eprintln!("Converted {} page(s)", files.len());
        }
        Err(e) => {
            eprintln!("Error: {}", e);
            process::exit(1);
        }
    }
}
