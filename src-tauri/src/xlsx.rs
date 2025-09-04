use calamine::{DataType, Reader, Xls, Xlsx, open_workbook};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::Path;

/// 学生基本信息结构
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Student {
    pub name: String,
    pub id_number: String,          // 身份证号
    pub student_id: Option<String>, // 学号
    pub class: Option<String>,      // 班级
    pub grade: Option<String>,      // 年级
    pub school: Option<String>,     // 学校
}

/// 困难学生信息结构
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DifficultStudent {
    pub name: String,
    pub id_number: String,
    pub difficulty_type: String,             // 困难类型
    pub source_file: String,                 // 来源文件
    pub extra_info: HashMap<String, String>, // 其他额外信息
}

/// 匹配结果结构
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MatchResult {
    pub student: Student,
    pub difficult_info: DifficultStudent,
}

/// Excel读取错误类型
#[derive(Debug)]
pub enum ExcelError {
    FileNotFound(String),
    ReadError(String),
    ParseError(String),
}

impl std::fmt::Display for ExcelError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match self {
            ExcelError::FileNotFound(path) => write!(f, "File not found: {}", path),
            ExcelError::ReadError(msg) => write!(f, "Read error: {}", msg),
            ExcelError::ParseError(msg) => write!(f, "Parse error: {}", msg),
        }
    }
}

impl std::error::Error for ExcelError {}

/// 清理和标准化身份证号
fn normalize_id_number(id: &str) -> String {
    let cleaned = id
        .trim()
        .replace(" ", "")
        .replace("\t", "")
        .replace("\n", "")
        .replace("\r", "")
        .to_uppercase();

    // 移除可能的 "G" 前缀
    if cleaned.starts_with("G") && cleaned.len() == 19 {
        cleaned[1..].to_string()
    } else {
        cleaned
    }
}

/// 读取学生信息表
pub fn read_student_info(file_path: &str) -> Result<Vec<Student>, Box<dyn std::error::Error>> {
    if !Path::new(file_path).exists() {
        return Err(Box::new(ExcelError::FileNotFound(file_path.to_string())));
    }

    let mut students = Vec::new();

    // 尝试读取 xlsx 格式
    if file_path.ends_with(".xlsx") {
        let mut workbook: Xlsx<_> = open_workbook(file_path)?;
        let range = workbook
            .worksheet_range_at(0)
            .ok_or("Cannot find first worksheet")??;

        // 假设第一行是表头，从第二行开始读取数据
        for (row_idx, row) in range.rows().enumerate() {
            if row_idx == 0 {
                continue; // 跳过表头
            }

            if row.len() >= 2 {
                // 调试：打印原始行数据

                let name = row
                    .first()
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string();

                // 尝试从第2列获取真实身份证号（去掉G前缀的）
                let id_number = row
                    .get(2)
                    .and_then(|v| v.as_string())
                    .or_else(|| row.get(1).and_then(|v| v.as_string()))
                    .unwrap_or_default()
                    .trim()
                    .to_string();

                if !name.is_empty() && !id_number.is_empty() {
                    let student = Student {
                        name,
                        id_number: normalize_id_number(&id_number),
                        student_id: row
                            .get(1)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        class: row
                            .get(11)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        grade: row
                            .get(10)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        school: row
                            .get(5)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                    };
                    students.push(student);
                }
            }
        }
    }

    Ok(students)
}

/// 读取困难类型表
pub fn read_difficult_type_table(
    file_path: &str,
    difficulty_type: &str,
) -> Result<Vec<DifficultStudent>, Box<dyn std::error::Error>> {
    if !Path::new(file_path).exists() {
        return Err(Box::new(ExcelError::FileNotFound(file_path.to_string())));
    }

    let mut difficult_students = Vec::new();
    let source_file = Path::new(file_path)
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string();

    if file_path.ends_with(".xlsx") {
        let mut workbook: Xlsx<_> = open_workbook(file_path)?;
        let range = workbook
            .worksheet_range_at(0)
            .ok_or("Cannot find first worksheet")??;

        // 读取表头以确定列位置
        let headers: Vec<String> = if let Some(header_row) = range.rows().next() {
            header_row
                .iter()
                .map(|cell| {
                    cell.as_string()
                        .unwrap_or_default()
                        .trim()
                        .to_string()
                        .to_lowercase()
                })
                .collect()
        } else {
            return Ok(difficult_students);
        };

        // // 查找姓名和身份证号的列位置
        // let name_col = find_column_index(&headers, &["姓名", "name", "名字"]);
        // let id_col = find_column_index(
        //     &headers,
        //     &["身份证号", "身份证", "id", "idcard", "证件号", "证件号码"],
        // );

        // 找到真正的数据开始行
        let mut data_start_row = 0;
        let mut real_headers: Vec<String> = Vec::new();

        // 扫描前几行找到真正的表头
        for (idx, row) in range.rows().enumerate().take(5) {
            let row_text: Vec<String> = row
                .iter()
                .map(|cell| cell.as_string().unwrap_or_default().trim().to_string())
                .collect();

            // 如果这行包含"姓名"或"证件号码"等关键词，认为是真正的表头
            if row_text.iter().any(|cell| {
                let cell_lower = cell.to_lowercase();
                cell_lower.contains("姓名")
                    || cell_lower.contains("证件号码")
                    || cell_lower.contains("身份证号")
            }) && row_text.iter().any(|cell| cell.contains("姓名"))
                && (row_text.iter().any(|cell| cell.contains("身份证"))
                    || row_text.iter().any(|cell| cell.contains("证件号码")))
            {
                real_headers = row_text;
                data_start_row = idx + 1;
                break;
            }
        }

        // 如果没找到合适的表头，使用原来的逻辑
        if real_headers.is_empty() {
            real_headers = headers.clone();
            data_start_row = 1;
        }

        // 重新查找姓名和身份证号的列位置
        let name_col = find_column_index(&real_headers, &["姓名", "name", "名字"]);
        let id_col = find_column_index(
            &real_headers,
            &["身份证号", "身份证", "id", "idcard", "证件号", "证件号码"],
        );

        for (row_idx, row) in range.rows().enumerate() {
            if row_idx < data_start_row {
                continue; // 跳过表头行
            }

            let name = if let Some(col) = name_col {
                row.get(col)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            } else {
                // 如果找不到姓名列，尝试使用第一列
                row.first()
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            };

            let id_number = if let Some(col) = id_col {
                row.get(col)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            } else {
                // 如果找不到身份证列，尝试使用第二列
                row.get(1)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            };

            if !name.is_empty()
                && !id_number.is_empty()
                && id_number.len() >= 15
                && !name.contains("序号")
                && !name.contains("姓名")
            {
                // 收集其他信息
                let mut extra_info = HashMap::new();
                for (col_idx, header) in headers.iter().enumerate() {
                    if !header.is_empty()
                        && col_idx != name_col.unwrap_or(0)
                        && col_idx != id_col.unwrap_or(1)
                        && let Some(cell_value) = row.get(col_idx).and_then(|v| v.as_string())
                            && !cell_value.trim().is_empty() {
                                extra_info.insert(header.clone(), cell_value.trim().to_string());
                            }
                }

                let difficult_student = DifficultStudent {
                    name,
                    id_number: normalize_id_number(&id_number),
                    difficulty_type: difficulty_type.to_string(),
                    source_file: source_file.clone(),
                    extra_info,
                };
                difficult_students.push(difficult_student);
            }
        }
    } else if file_path.ends_with(".xls") {
        // 处理旧版 Excel 格式
        let mut workbook: Xls<_> = open_workbook(file_path)?;
        let range = workbook
            .worksheet_range_at(0)
            .ok_or("Cannot find first worksheet")??;

        // 读取表头
        let headers: Vec<String> = if let Some(header_row) = range.rows().next() {
            header_row
                .iter()
                .map(|cell| {
                    cell.as_string()
                        .unwrap_or_default()
                        .trim()
                        .to_string()
                        .to_lowercase()
                })
                .collect()
        } else {
            return Ok(difficult_students);
        };

        let name_col = find_column_index(&headers, &["姓名", "name", "名字"]);
        let id_col = find_column_index(
            &headers,
            &["身份证号", "身份证", "id", "idcard", "证件号", "证件号码"],
        );

        for (row_idx, row) in range.rows().enumerate() {
            if row_idx == 0 {
                continue;
            }

            let name = if let Some(col) = name_col {
                row.get(col)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            } else {
                row.first()
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            };

            let id_number = if let Some(col) = id_col {
                row.get(col)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            } else {
                row.get(1)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            };

            if !name.is_empty()
                && !id_number.is_empty()
                && id_number.len() >= 15
                && !name.contains("序号")
                && !name.contains("姓名")
            {
                let mut extra_info = HashMap::new();
                for (col_idx, header) in headers.iter().enumerate() {
                    if !header.is_empty()
                        && col_idx != name_col.unwrap_or(0)
                        && col_idx != id_col.unwrap_or(1)
                        && let Some(cell_value) = row.get(col_idx).and_then(|v| v.as_string())
                            && !cell_value.trim().is_empty() {
                                extra_info.insert(header.clone(), cell_value.trim().to_string());
                            }
                }

                let difficult_student = DifficultStudent {
                    name,
                    id_number: normalize_id_number(&id_number),
                    difficulty_type: difficulty_type.to_string(),
                    source_file: source_file.clone(),
                    extra_info,
                };
                difficult_students.push(difficult_student);
            }
        }
    }

    Ok(difficult_students)
}

/// 查找列索引的辅助函数
fn find_column_index(headers: &[String], possible_names: &[&str]) -> Option<usize> {
    for (idx, header) in headers.iter().enumerate() {
        for name in possible_names {
            if header.contains(name) {
                return Some(idx);
            }
        }
    }
    None
}

/// 匹配学生信息和困难类型信息
pub fn match_students_with_difficulty(
    students: &[Student],
    difficult_students: &[DifficultStudent],
) -> Vec<MatchResult> {
    let mut results = Vec::new();

    // 创建学生身份证号的哈希映射以提高查询效率
    let student_map: HashMap<String, &Student> =
        students.iter().map(|s| (s.id_number.clone(), s)).collect();

    for difficult_student in difficult_students {
        if let Some(student) = student_map.get(&difficult_student.id_number) {
            results.push(MatchResult {
                student: (*student).clone(),
                difficult_info: difficult_student.clone(),
            });
        }
    }

    results
}

/// 处理所有困难类型表并匹配学生信息
pub fn process_all_difficulty_types(
    students: &[Student],
    data_dir: &str,
) -> Result<Vec<MatchResult>, Box<dyn std::error::Error>> {
    let difficulty_types = vec![
        (
            "1-脱贫户信息_20250902-继续享受政策.xlsx",
            "脱贫户(继续享受政策)",
        ),
        ("2-2025.9.3持证残疾人名单.xlsx", "持证残疾人"),
        ("3-02025.9.1-2025年9月份农村低保备案表.xls", "农村低保"),
        (
            "4-2025.9.1-2025年9月份城乡特困人员备案表.xlsx",
            "城乡特困人员",
        ),
        ("5-2025.9.1-2025年9月份城镇低保全部备案表.xls", "城镇低保"),
        (
            "6-防止返贫致贫监测对象信息_20250902-风险未消除608人.xlsx",
            "防返贫监测对象(风险未消除)",
        ),
        (
            "7-2025.9.1-2025年9月份孤儿及事实无人抚养儿童发放花名册.xls",
            "孤儿及事实无人抚养儿童",
        ),
        ("8-2025.9.2-2025年9月低收入人口花名册.xlsx", "低收入人口"),
        (
            "9-脱贫户信息_20250902-脱贫不享受政策.xlsx",
            "脱贫户(不享受政策)",
        ),
        (
            "10-防止返贫致贫监测对象信息_20250902-风险已消除2928人.xlsx",
            "防返贫监测对象(风险已消除)",
        ),
    ];

    let mut all_results = Vec::new();

    for (filename, difficulty_type) in difficulty_types {
        let file_path = format!(
            "{}/2025年秋季学期困难类型查询专用表及最新的本县户籍特殊困难群体人员信息9.3/{}",
            data_dir, filename
        );

        match read_difficult_type_table(&file_path, difficulty_type) {
            Ok(difficult_students) => {
                let matches = match_students_with_difficulty(students, &difficult_students);
                all_results.extend(matches);
            }
            Err(_) => {
                // 忽略无法处理的文件
            }
        }
    }

    Ok(all_results)
}

/// 公开的API函数：处理学生信息匹配
pub fn find_difficulty_students(
    data_dir: &str,
) -> Result<Vec<MatchResult>, Box<dyn std::error::Error>> {
    // 读取学生信息表
    let student_info_path = format!("{}/20250904在校生信息.xlsx", data_dir);
    let students = read_student_info(&student_info_path)?;

    // 处理所有困难类型表
    let results = process_all_difficulty_types(&students, data_dir)?;

    Ok(results)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_id_number() {
        assert_eq!(
            normalize_id_number("  123456789012345678  "),
            "123456789012345678"
        );
        assert_eq!(
            normalize_id_number("123\t456\n789\r012345678"),
            "123456789012345678"
        );
    }

    #[test]
    fn test_find_column_index() {
        let headers = vec![
            "序号".to_string(),
            "姓名".to_string(),
            "身份证号".to_string(),
        ];
        assert_eq!(find_column_index(&headers, &["姓名"]), Some(1));
        assert_eq!(
            find_column_index(&headers, &["身份证号", "身份证"]),
            Some(2)
        );
        assert_eq!(find_column_index(&headers, &["不存在的列"]), None);
    }
}
