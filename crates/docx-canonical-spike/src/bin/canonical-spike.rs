use std::io::{self, BufRead, Write};

fn main() -> io::Result<()> {
    let stdin = io::stdin();
    let mut stdout = io::BufWriter::new(io::stdout().lock());
    for line in stdin.lock().lines() {
        let line = line?;
        let response = match serde_json::from_str::<serde_json::Value>(&line) {
            Ok(request) if request.get("harness").is_some() => {
                docx_canonical_spike::harness::apply_request(&request["harness"])
            }
            Ok(request) => match request.get("benchmark") {
                Some(benchmark) => docx_canonical_spike::benchmark::run(benchmark),
                None => docx_canonical_spike::apply_request(&request),
            },
            Err(error) => {
                serde_json::json!({"status":"transportError","message":error.to_string()})
            }
        };
        serde_json::to_writer(&mut stdout, &response)?;
        writeln!(stdout)?;
        stdout.flush()?;
    }
    Ok(())
}
