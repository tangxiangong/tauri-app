pub mod command;
pub mod xlsx;

use command::*;

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            find_students_by_difficulty,
            get_students_match_statistics,
            validate_uploaded_file,
            get_difficulty_type_options,
            export_matches_to_excel,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
