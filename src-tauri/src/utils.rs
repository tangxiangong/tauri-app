pub fn format_bytes(bytes: u64) -> String {
    let units = ["B", "KB", "MB", "GB"];
    let mut bytes = bytes as f64;
    let mut unit = 0;

    while bytes >= 1024.0 && unit < units.len() - 1 {
        bytes /= 1024.0;
        unit += 1;
    }

    format!("{:.2} {}", bytes, units[unit])
}
