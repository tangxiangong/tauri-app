#[tauri::command]
pub fn is_supported() -> bool {
    crate::core::is_supported()
}

#[tauri::command]
pub fn get_memory() -> crate::core::Memory {
    crate::core::get_memory()
}
