use calamine::{DataType, Reader, Xls, Xlsx, open_workbook};
use serde::{Deserialize, Serialize};
use std::{collections::HashMap, path::Path};

/// 困难类型枚举
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum DifficultyType {
    #[serde(rename = "脱贫户(继续享受政策)")]
    PovertyAlleviatedContinuePolicy,
    #[serde(rename = "脱贫户(不享受政策)")]
    PovertyAlleviatedNoPolicy,
    #[serde(rename = "持证残疾人")]
    DisabledWithCertificate,
    #[serde(rename = "农村低保")]
    RuralMinimumLiving,
    #[serde(rename = "城镇低保")]
    UrbanMinimumLiving,
    #[serde(rename = "城乡特困")]
    RuralSpecialDifficulty,
    #[serde(rename = "防返贫监测对象(风险未消除)")]
    AntiPovertyMonitoringRiskNotEliminated,
    #[serde(rename = "防返贫监测对象(风险已消除)")]
    AntiPovertyMonitoringRiskEliminated,
    #[serde(rename = "孤儿及事实无人抚养儿童")]
    OrphansAndFactuallyUnsupportedChildren,
    #[serde(rename = "低收入人口")]
    LowIncomePopulation,
}

impl std::fmt::Display for DifficultyType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::PovertyAlleviatedContinuePolicy => write!(f, "脱贫户(继续享受政策)"),
            Self::PovertyAlleviatedNoPolicy => write!(f, "脱贫户(不享受政策)"),
            Self::DisabledWithCertificate => write!(f, "持证残疾人"),
            Self::RuralMinimumLiving => write!(f, "农村低保"),
            Self::UrbanMinimumLiving => write!(f, "城镇低保"),
            Self::RuralSpecialDifficulty => write!(f, "城乡特困"),
            Self::AntiPovertyMonitoringRiskNotEliminated => write!(f, "防返贫监测对象(风险未消除)"),
            Self::AntiPovertyMonitoringRiskEliminated => write!(f, "防返贫监测对象(风险已消除)"),
            Self::OrphansAndFactuallyUnsupportedChildren => write!(f, "孤儿及事实无人抚养儿童"),
            Self::LowIncomePopulation => write!(f, "低收入人口"),
        }
    }
}

impl DifficultyType {
    /// 获取所有困难类型
    pub fn all() -> Vec<Self> {
        vec![
            Self::PovertyAlleviatedContinuePolicy,
            Self::PovertyAlleviatedNoPolicy,
            Self::DisabledWithCertificate,
            Self::RuralMinimumLiving,
            Self::UrbanMinimumLiving,
            Self::RuralSpecialDifficulty,
            Self::AntiPovertyMonitoringRiskNotEliminated,
            Self::AntiPovertyMonitoringRiskEliminated,
            Self::OrphansAndFactuallyUnsupportedChildren,
            Self::LowIncomePopulation,
        ]
    }

    /// 根据困难类型获取列配置 (姓名列索引, 身份证列索引, 数据开始行, 工作表索引)
    pub fn get_column_config(&self) -> (Option<usize>, Option<usize>, usize, usize) {
        match self {
            Self::PovertyAlleviatedContinuePolicy | Self::PovertyAlleviatedNoPolicy => {
                (None, Some(7), 1, 0) // H列(索引7), 第1个工作表
            }
            Self::DisabledWithCertificate => (Some(0), Some(1), 1, 0), // A列姓名，B列身份证号, 第1个工作表
            Self::RuralMinimumLiving => (Some(4), Some(6), 1, 1), // E列姓名，G列身份证号, 第2个工作表
            Self::UrbanMinimumLiving => (None, Some(6), 1, 0),    // G列(索引6), 第1个工作表
            Self::RuralSpecialDifficulty => (Some(0), Some(5), 1, 1), // A列姓名，F列身份证号, 第2个工作表
            Self::AntiPovertyMonitoringRiskNotEliminated
            | Self::AntiPovertyMonitoringRiskEliminated => (None, Some(11), 1, 0), // L列(索引11), 第1个工作表
            Self::OrphansAndFactuallyUnsupportedChildren => (None, Some(2), 1, 0), // C列(索引2), 第1个工作表
            Self::LowIncomePopulation => (None, Some(3), 1, 0), // D列(索引3), 第1个工作表
        }
    }
}

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
    pub id_number: String,                   // 身份证号
    pub difficulty_type: DifficultyType,     // 困难类型
    pub source_file: String,                 // 来源文件
    pub extra_info: HashMap<String, String>, // 其他信息
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
    id.trim()
        .replace(" ", "")
        .replace("\t", "")
        .replace("\n", "")
        .replace("\r", "")
        .to_uppercase()
}

/// 读取学生信息表
pub fn read_student_info(file_path: &str) -> Result<Vec<Student>, Box<dyn std::error::Error>> {
    if !Path::new(file_path).exists() {
        return Err(Box::new(ExcelError::FileNotFound(file_path.to_string())));
    }

    let mut students = Vec::new();

    // 尝试读取 xlsx 格式
    // 处理 xlsx 格式
    if file_path.ends_with(".xlsx") {
        let mut workbook: Xlsx<_> = open_workbook(file_path)?;
        let range = workbook
            .worksheet_range_at(0)
            .ok_or("Cannot find first worksheet")??;

        for (row_idx, row) in range.rows().enumerate() {
            if row_idx == 0 {
                continue; // 跳过表头
            }

            if row.len() >= 3 {
                // A列：学生姓名
                let name = row
                    .first()
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string();

                // C列：身份证件号
                let id_number = row
                    .get(2)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string();

                if !name.is_empty() && !id_number.is_empty() {
                    let student = Student {
                        name,
                        id_number: normalize_id_number(&id_number),
                        // B列：全国学籍号
                        student_id: row
                            .get(1)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // L列：班级
                        class: row
                            .get(11)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // K列：年级
                        grade: row
                            .get(10)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // F列：学校名称
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
    difficulty_type: DifficultyType,
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

    // 根据困难类型确定列位置
    let (name_col, id_col, data_start_row, target_worksheet_index) =
        difficulty_type.get_column_config();

    // 处理 .xlsx 格式
    if file_path.ends_with(".xlsx") {
        let mut workbook: Xlsx<_> = open_workbook(file_path)?;
        let range = workbook
            .worksheet_range_at(target_worksheet_index)
            .ok_or(format!(
                "Cannot find worksheet at index {}",
                target_worksheet_index
            ))??;

        // 跳过可能的标题行，查找真正的数据开始行
        let mut actual_start_row = data_start_row;
        for (row_idx, row) in range.rows().enumerate().skip(data_start_row) {
            for cell in row {
                if let Some(cell_str) = cell.as_string() {
                    let cell_str = cell_str.trim();
                    // 如果找到数字，认为是数据开始
                    if cell_str.len() >= 10 && cell_str.chars().any(|c| c.is_ascii_digit()) {
                        actual_start_row = row_idx;
                        break;
                    }
                }
            }
            if actual_start_row != data_start_row {
                break;
            }
        }

        for (row_idx, row) in range.rows().enumerate() {
            if row_idx < actual_start_row {
                continue; // 跳过标题行和表头
            }

            // 获取姓名（如果有的话）
            let name = if let Some(col) = name_col {
                row.get(col)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            } else {
                // 对于没有姓名列的表，使用"未知"
                "未知".to_string()
            };

            // 获取身份证号
            let id_number = if let Some(col) = id_col {
                row.get(col)
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string()
            } else {
                continue; // 没有身份证号列则跳过
            };

            // 只要身份证号不为空就添加记录
            if !id_number.is_empty() {
                let difficult_student = DifficultStudent {
                    name,
                    id_number: normalize_id_number(&id_number),
                    difficulty_type: difficulty_type.clone(),
                    source_file: source_file.clone(),
                    extra_info: HashMap::new(),
                };
                difficult_students.push(difficult_student);
            }
        }
    } else if file_path.ends_with(".xls") {
        // 处理旧版 Excel 格式
        let mut workbook: Xls<_> = open_workbook(file_path)?;
        let range = workbook
            .worksheet_range_at(target_worksheet_index)
            .ok_or(format!(
                "Cannot find worksheet at index {}",
                target_worksheet_index
            ))??;

        for (row_idx, row) in range.rows().enumerate() {
            if row_idx < data_start_row {
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

            if !id_number.is_empty() {
                let difficult_student = DifficultStudent {
                    name,
                    id_number: normalize_id_number(&id_number),
                    difficulty_type: difficulty_type.clone(),
                    source_file: source_file.clone(),
                    extra_info: HashMap::new(),
                };
                difficult_students.push(difficult_student);
            }
        }
    }

    Ok(difficult_students)
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
