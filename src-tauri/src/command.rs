use crate::xlsx::{
    MatchResult, find_difficulty_students, match_students_with_difficulty,
    read_difficult_type_table, read_student_info,
};
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
    pub difficulty_type_counts: std::collections::HashMap<String, usize>,
}

/// 处理学生信息匹配的主命令
#[tauri::command]
pub async fn find_difficult_students(data_dir: String) -> CommandResult<Vec<MatchResult>> {
    match find_difficulty_students(&data_dir) {
        Ok(results) => CommandResult::success(results),
        Err(e) => CommandResult::error(format!("匹配失败: {}", e)),
    }
}

/// 获取匹配结果的统计信息
#[tauri::command]
pub async fn get_match_statistics(data_dir: String) -> CommandResult<MatchStatistics> {
    match find_difficulty_students(&data_dir) {
        Ok(results) => {
            let total_matches = results.len();

            // 统计各困难类型的人数
            let mut difficulty_type_counts = std::collections::HashMap::new();
            for result in &results {
                *difficulty_type_counts
                    .entry(result.difficult_info.difficulty_type.clone())
                    .or_insert(0) += 1;
            }

            let statistics = MatchStatistics {
                total_students: 0, // 这里可以根据需要添加总学生数的计算
                total_matches,
                difficulty_type_counts,
            };

            CommandResult::success(statistics)
        }
        Err(e) => CommandResult::error(format!("获取统计信息失败: {}", e)),
    }
}

/// 验证数据目录是否存在且包含必要文件
#[tauri::command]
pub async fn validate_data_directory(data_dir: String) -> CommandResult<bool> {
    let data_path = PathBuf::from(&data_dir);

    if !data_path.exists() {
        return CommandResult::error("数据目录不存在".to_string());
    }

    // 检查学生信息表是否存在
    let student_info_file = data_path.join("20250904在校生信息.xlsx");
    if !student_info_file.exists() {
        return CommandResult::error("学生信息表文件不存在".to_string());
    }

    // 检查困难类型表目录是否存在
    let difficulty_dir =
        data_path.join("2025年秋季学期困难类型查询专用表及最新的本县户籍特殊困难群体人员信息9.3");
    if !difficulty_dir.exists() {
        return CommandResult::error("困难类型查询表目录不存在".to_string());
    }

    // 检查关键的困难类型表文件是否存在
    let required_files = vec![
        "1-脱贫户信息_20250902-继续享受政策.xlsx",
        "2-2025.9.3持证残疾人名单.xlsx",
        "3-02025.9.1-2025年9月份农村低保备案表.xls",
        "4-2025.9.1-2025年9月份城乡特困人员备案表.xlsx",
        "5-2025.9.1-2025年9月份城镇低保全部备案表.xls",
    ];

    let mut missing_files = Vec::new();
    for file in &required_files {
        let file_path = difficulty_dir.join(file);
        if !file_path.exists() {
            missing_files.push(file.to_string());
        }
    }

    if !missing_files.is_empty() {
        return CommandResult::error(format!("缺少以下文件: {}", missing_files.join(", ")));
    }

    CommandResult::success(true)
}

/// 根据困难类型过滤匹配结果
#[tauri::command]
pub async fn filter_by_difficulty_type(
    data_dir: String,
    difficulty_type: String,
) -> CommandResult<Vec<MatchResult>> {
    match find_difficulty_students(&data_dir) {
        Ok(results) => {
            let filtered: Vec<MatchResult> = results
                .into_iter()
                .filter(|r| r.difficult_info.difficulty_type == difficulty_type)
                .collect();

            CommandResult::success(filtered)
        }
        Err(e) => CommandResult::error(format!("过滤失败: {}", e)),
    }
}

/// 获取所有困难类型列表
#[tauri::command]
pub async fn get_difficulty_types(data_dir: String) -> CommandResult<Vec<String>> {
    match find_difficulty_students(&data_dir) {
        Ok(results) => {
            let mut types: std::collections::HashSet<String> = std::collections::HashSet::new();
            for result in results {
                types.insert(result.difficult_info.difficulty_type);
            }

            let mut sorted_types: Vec<String> = types.into_iter().collect();
            sorted_types.sort();

            CommandResult::success(sorted_types)
        }
        Err(e) => CommandResult::error(format!("获取困难类型失败: {}", e)),
    }
}

/// 根据学生姓名搜索匹配结果
#[tauri::command]
pub async fn search_by_student_name(
    data_dir: String,
    name: String,
) -> CommandResult<Vec<MatchResult>> {
    match find_difficulty_students(&data_dir) {
        Ok(results) => {
            let filtered: Vec<MatchResult> = results
                .into_iter()
                .filter(|r| r.student.name.contains(&name))
                .collect();

            CommandResult::success(filtered)
        }
        Err(e) => CommandResult::error(format!("搜索失败: {}", e)),
    }
}

/// 导出匹配结果为 JSON 格式
#[tauri::command]
pub async fn export_results_as_json(
    data_dir: String,
    output_path: String,
) -> CommandResult<String> {
    match find_difficulty_students(&data_dir) {
        Ok(results) => match serde_json::to_string_pretty(&results) {
            Ok(json_str) => match std::fs::write(&output_path, &json_str) {
                Ok(_) => CommandResult::success(format!("成功导出到: {}", output_path)),
                Err(e) => CommandResult::error(format!("写入文件失败: {}", e)),
            },
            Err(e) => CommandResult::error(format!("序列化为 JSON 失败: {}", e)),
        },
        Err(e) => CommandResult::error(format!("获取数据失败: {}", e)),
    }
}

/// 检查身份证号格式是否有效
pub fn validate_id_number(id: &str) -> bool {
    // 基本格式检查：15位或18位数字，最后一位可能是X
    if id.len() != 15 && id.len() != 18 {
        return false;
    }

    // 检查前面的位数是否都是数字
    let chars: Vec<char> = id.chars().collect();
    for &ch in &chars[..chars.len() - 1] {
        if !ch.is_ascii_digit() {
            return false;
        }
    }

    // 最后一位可以是数字或X
    let last_char = chars[chars.len() - 1];
    last_char.is_ascii_digit() || last_char == 'X' || last_char == 'x'
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_validate_id_number() {
        assert!(validate_id_number("123456789012345678"));
        assert!(validate_id_number("12345678901234567X"));
        assert!(validate_id_number("123456789012345"));
        assert!(!validate_id_number("12345678901234567"));
        assert!(!validate_id_number("1234567890123456789"));
        assert!(!validate_id_number("12345678901234567A"));
    }

    #[test]
    fn test_command_result() {
        let success_result: CommandResult<i32> = CommandResult::success(42);
        assert!(success_result.success);
        assert_eq!(success_result.data, Some(42));
        assert!(success_result.error.is_none());

        let error_result: CommandResult<i32> = CommandResult::error("test error".to_string());
        assert!(!error_result.success);
        assert!(error_result.data.is_none());
        assert_eq!(error_result.error, Some("test error".to_string()));
    }
}

/// 处理上传的学生文件和困难类型文件
#[tauri::command]
pub async fn process_uploaded_files(
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

    // 读取困难类型表
    let difficult_students =
        match read_difficult_type_table(&difficulty_file_path, &difficulty_type) {
            Ok(difficult_students) => difficult_students,
            Err(e) => {
                return CommandResult::error(format!("读取困难类型文件失败: {}", e));
            }
        };

    // 匹配学生信息
    let matches = match_students_with_difficulty(&students, &difficult_students);

    CommandResult::success(matches)
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
    let options = vec![
        DifficultyTypeOption {
            label: "脱贫户(继续享受政策)".to_string(),
            value: "脱贫户(继续享受政策)".to_string(),
        },
        DifficultyTypeOption {
            label: "脱贫户(不享受政策)".to_string(),
            value: "脱贫户(不享受政策)".to_string(),
        },
        DifficultyTypeOption {
            label: "持证残疾人".to_string(),
            value: "持证残疾人".to_string(),
        },
        DifficultyTypeOption {
            label: "农村低保".to_string(),
            value: "农村低保".to_string(),
        },
        DifficultyTypeOption {
            label: "城镇低保".to_string(),
            value: "城镇低保".to_string(),
        },
        DifficultyTypeOption {
            label: "城乡特困人员".to_string(),
            value: "城乡特困人员".to_string(),
        },
        DifficultyTypeOption {
            label: "防返贫监测对象(风险未消除)".to_string(),
            value: "防返贫监测对象(风险未消除)".to_string(),
        },
        DifficultyTypeOption {
            label: "防返贫监测对象(风险已消除)".to_string(),
            value: "防返贫监测对象(风险已消除)".to_string(),
        },
        DifficultyTypeOption {
            label: "孤儿及事实无人抚养儿童".to_string(),
            value: "孤儿及事实无人抚养儿童".to_string(),
        },
        DifficultyTypeOption {
            label: "低收入人口".to_string(),
            value: "低收入人口".to_string(),
        },
    ];

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
