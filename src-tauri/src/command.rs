use crate::xlsx::{
    DifficultyType, MatchResult, match_students_with_difficulty, read_difficult_type_table,
    read_student_info,
};
use rust_xlsxwriter::{Format, Workbook};
use serde::{Deserialize, Serialize};
use std::path::PathBuf;

/// 命令执行结果
#[derive(Debug, Serialize, Deserialize)]
pub struct CommandResult<T> {
    pub success: bool,
    pub data: Option<T>,
    pub error: Option<String>,
}

impl<T> CommandResult<T> {
    pub fn success(data: T) -> Self {
        Self {
            success: true,
            data: Some(data),
            error: None,
        }
    }

    pub fn error(error: String) -> Self {
        Self {
            success: false,
            data: None,
            error: Some(error),
        }
    }
}

/// 匹配结果统计信息
#[derive(Debug, Serialize, Deserialize)]
pub struct MatchStatistics {
    pub total_students: usize,
    pub total_matches: usize,
    pub difficulty_type_counts: std::collections::HashMap<DifficultyType, usize>,
}

/// 根据困难类型查找学生信息
#[tauri::command]
pub async fn find_students_by_difficulty(
    student_file_path: String,
    difficulty_file_path: String,
    difficulty_type: String,
) -> CommandResult<Vec<MatchResult>> {
    // 读取学生信息
    let students = match read_student_info(&student_file_path) {
        Ok(students) => students,
        Err(e) => {
            return CommandResult::error(format!("读取学生文件失败: {}", e));
        }
    };

    // 解析困难类型枚举
    let difficulty_json = format!(r#""{}""#, difficulty_type);
    let difficulty_enum: DifficultyType = match serde_json::from_str(&difficulty_json) {
        Ok(enum_val) => enum_val,
        Err(_) => {
            return CommandResult::error(format!("未知的困难类型: {}", difficulty_type));
        }
    };

    // 读取困难类型表
    let difficult_students = match read_difficult_type_table(&difficulty_file_path, difficulty_enum)
    {
        Ok(difficult_students) => difficult_students,
        Err(e) => {
            return CommandResult::error(format!("读取困难类型文件失败: {}", e));
        }
    };

    // 匹配学生信息
    let matches = match_students_with_difficulty(&students, &difficult_students);

    CommandResult::success(matches)
}

/// 获取匹配结果统计信息
#[tauri::command]
pub async fn get_students_match_statistics(
    student_file_path: String,
    difficulty_file_path: String,
    difficulty_type: String,
) -> CommandResult<MatchStatistics> {
    // 复用查找逻辑
    let result =
        find_students_by_difficulty(student_file_path, difficulty_file_path, difficulty_type).await;

    match result {
        CommandResult {
            success: true,
            data: Some(matches),
            ..
        } => {
            let mut difficulty_type_counts = std::collections::HashMap::new();

            for match_result in &matches {
                *difficulty_type_counts
                    .entry(match_result.difficult_info.difficulty_type)
                    .or_insert(0) += 1;
            }

            let statistics = MatchStatistics {
                total_students: matches.len(),
                total_matches: matches.len(),
                difficulty_type_counts,
            };

            CommandResult::success(statistics)
        }
        CommandResult {
            error: Some(err), ..
        } => CommandResult::error(err),
        _ => CommandResult::error("获取统计信息失败".to_string()),
    }
}

/// 验证上传的文件
#[tauri::command]
pub async fn validate_uploaded_file(file_path: String) -> CommandResult<FileInfo> {
    let path = PathBuf::from(&file_path);

    if !path.exists() {
        return CommandResult::error("文件不存在".to_string());
    }

    let file_name = path
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string();

    let file_extension = path
        .extension()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string()
        .to_lowercase();

    if file_extension != "xlsx" && file_extension != "xls" {
        return CommandResult::error("仅支持 Excel 文件 (.xlsx 或 .xls)".to_string());
    }

    let file_size = match std::fs::metadata(&path) {
        Ok(metadata) => metadata.len(),
        Err(e) => return CommandResult::error(format!("无法读取文件信息: {}", e)),
    };

    let file_info = FileInfo {
        name: file_name,
        path: file_path,
        size: file_size,
        extension: file_extension,
    };

    CommandResult::success(file_info)
}

/// 获取困难类型选项
#[tauri::command]
pub async fn get_difficulty_type_options() -> CommandResult<Vec<DifficultyTypeOption>> {
    let options = DifficultyType::all()
        .into_iter()
        .map(|difficulty_type| DifficultyTypeOption {
            label: difficulty_type.to_string(),
            value: difficulty_type.to_string(),
        })
        .collect();

    CommandResult::success(options)
}

/// 文件信息结构
#[derive(Debug, Serialize, Deserialize)]
pub struct FileInfo {
    pub name: String,
    pub path: String,
    pub size: u64,
    pub extension: String,
}

/// 困难类型选项
#[derive(Debug, Serialize, Deserialize)]
pub struct DifficultyTypeOption {
    pub label: String,
    pub value: String,
}

/// 导出匹配结果到 Excel 文件
#[tauri::command]
pub async fn export_matches_to_excel(
    matches: Vec<MatchResult>,
    output_path: String,
) -> CommandResult<String> {
    match create_excel_report(&matches, &output_path) {
        Ok(_) => CommandResult::success(output_path),
        Err(e) => CommandResult::error(format!("导出 Excel 失败: {}", e)),
    }
}

/// 创建 Excel 报告
fn create_excel_report(
    matches: &[MatchResult],
    output_path: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // 设置标题格式
    let header_format = Format::new()
        .set_bold()
        .set_background_color("#4CAF50")
        .set_font_color("#FFFFFF")
        .set_align(rust_xlsxwriter::FormatAlign::Center);

    // 写入标题行
    let headers = [
        "序号",
        "学生姓名",
        "身份证号",
        "学号",
        "班级",
        "年级",
        "学校",
        "困难类型",
    ];

    for (col, header) in headers.iter().enumerate() {
        worksheet.write_with_format(0, col as u16, *header, &header_format)?;
    }

    // 设置数据格式
    let data_format = Format::new().set_align(rust_xlsxwriter::FormatAlign::Left);
    let number_format = Format::new().set_align(rust_xlsxwriter::FormatAlign::Center);

    // 写入数据
    for (row, match_result) in matches.iter().enumerate() {
        let row = row + 1; // 跳过标题行

        worksheet.write_with_format(row as u32, 0, row as u32, &number_format)?;
        worksheet.write_with_format(row as u32, 1, &match_result.student.name, &data_format)?;
        worksheet.write_with_format(
            row as u32,
            2,
            &match_result.student.id_number,
            &data_format,
        )?;
        worksheet.write_with_format(
            row as u32,
            3,
            match_result.student.student_id.as_deref().unwrap_or(""),
            &data_format,
        )?;
        worksheet.write_with_format(
            row as u32,
            4,
            match_result.student.class.as_deref().unwrap_or(""),
            &data_format,
        )?;
        worksheet.write_with_format(
            row as u32,
            5,
            match_result.student.grade.as_deref().unwrap_or(""),
            &data_format,
        )?;
        worksheet.write_with_format(
            row as u32,
            6,
            match_result.student.school.as_deref().unwrap_or(""),
            &data_format,
        )?;
        worksheet.write_with_format(
            row as u32,
            7,
            match_result.difficult_info.difficulty_type.to_string(),
            &data_format,
        )?;
    }

    // 设置列宽
    worksheet.set_column_width(0, 6.0)?; // 序号
    worksheet.set_column_width(1, 12.0)?; // 姓名
    worksheet.set_column_width(2, 20.0)?; // 身份证号
    worksheet.set_column_width(3, 15.0)?; // 学号
    worksheet.set_column_width(4, 12.0)?; // 班级
    worksheet.set_column_width(5, 8.0)?; // 年级
    worksheet.set_column_width(6, 20.0)?; // 学校
    worksheet.set_column_width(7, 18.0)?; // 困难类型

    // 添加统计信息工作表
    let stats_worksheet = workbook.add_worksheet();
    stats_worksheet.set_name("统计信息")?;

    // 写入统计信息标题
    stats_worksheet.write_with_format(0, 0, "统计项目", &header_format)?;
    stats_worksheet.write_with_format(0, 1, "数量", &header_format)?;

    // 计算统计信息
    let mut difficulty_counts = std::collections::HashMap::new();
    for match_result in matches {
        *difficulty_counts
            .entry(match_result.difficult_info.difficulty_type.to_string())
            .or_insert(0) += 1;
    }

    let mut row = 1;
    stats_worksheet.write_with_format(row as u32, 0, "总匹配数量", &data_format)?;
    stats_worksheet.write_with_format(row as u32, 1, matches.len() as u32, &number_format)?;
    row += 1;

    stats_worksheet.write_with_format(row as u32, 0, "按困难类型分布:", &data_format)?;
    row += 1;

    for (difficulty_type, count) in difficulty_counts.iter() {
        stats_worksheet.write_with_format(row as u32, 0, difficulty_type, &data_format)?;
        stats_worksheet.write_with_format(row as u32, 1, *count as u32, &number_format)?;
        row += 1;
    }

    // 设置统计表列宽
    stats_worksheet.set_column_width(0, 25.0)?;
    stats_worksheet.set_column_width(1, 10.0)?;

    workbook.save(output_path)?;
    Ok(())
}
