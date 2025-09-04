pub mod command;
pub mod core;
pub mod utils;

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            command::is_supported,
            command::get_memory
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
