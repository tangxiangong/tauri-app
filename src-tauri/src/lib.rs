pub mod command;
pub mod xlsx;

use command::{
    export_results_as_json, filter_by_difficulty_type, find_difficult_students,
    get_difficulty_type_options, get_difficulty_types, get_match_statistics,
    process_uploaded_files, search_by_student_name, validate_data_directory,
    validate_uploaded_file,
};

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            find_difficult_students,
            get_match_statistics,
            validate_data_directory,
            filter_by_difficulty_type,
            get_difficulty_types,
            search_by_student_name,
            export_results_as_json,
            process_uploaded_files,
            validate_uploaded_file,
            get_difficulty_type_options,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
