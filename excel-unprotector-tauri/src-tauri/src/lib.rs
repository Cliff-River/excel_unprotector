pub mod excel;

use excel::unprotect;

#[tauri::command]
fn unprotect_excel(file_name: String, data: Vec<u8>) -> Result<tauri::ipc::Response, String> {
    if file_name.is_empty() || !file_name.to_lowercase().ends_with(".xlsx") {
        return Err("只支持 .xlsx 格式的 Excel 文件".to_string());
    }
    match unprotect(&data) {
        Ok(bytes) => Ok(tauri::ipc::Response::new(bytes)),
        Err(err) => Err(err.to_string()),
    }
}

#[tauri::command]
fn read_dropped_file(path: String) -> Result<tauri::ipc::Response, String> {
    if !path.to_lowercase().ends_with(".xlsx") {
        return Err("只支持 .xlsx 格式的 Excel 文件".to_string());
    }
    match std::fs::read(&path) {
        Ok(bytes) => Ok(tauri::ipc::Response::new(bytes)),
        Err(_) => Err("无法读取拖拽的文件，请重试".to_string()),
    }
}

#[tauri::command]
async fn save_file(
    app: tauri::AppHandle,
    default_file_name: String,
    data: Vec<u8>,
) -> Result<Option<String>, String> {
    use tauri_plugin_dialog::DialogExt;

    let path = app
        .dialog()
        .file()
        .set_file_name(&default_file_name)
        .add_filter("Excel 文件", &["xlsx"])
        .blocking_save_file();

    let Some(file_path) = path else {
        return Ok(None);
    };

    let path_buf = file_path
        .into_path()
        .map_err(|_| "保存文件失败，请重试".to_string())?;

    std::fs::write(&path_buf, &data).map_err(|_| "保存文件失败，请重试".to_string())?;
    Ok(Some(path_buf.to_string_lossy().into_owned()))
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            unprotect_excel,
            save_file,
            read_dropped_file
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
