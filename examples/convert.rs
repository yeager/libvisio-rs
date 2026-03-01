//! Example: Convert a Visio file to SVG.

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <input.vsdx> [output_dir]", args[0]);
        std::process::exit(1);
    }

    let input = &args[1];
    let output_dir = args.get(2).map(|s| s.as_str()).unwrap_or(".");

    match libvisio_rs::convert(input, Some(output_dir), None) {
        Ok(files) => {
            for f in &files {
                println!("Created: {}", f);
            }
        }
        Err(e) => {
            eprintln!("Error: {}", e);
            std::process::exit(1);
        }
    }
}
