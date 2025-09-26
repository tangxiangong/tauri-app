use calamine::{DataType, Reader, Xls, XlsError, Xlsx, XlsxError, open_workbook};
use serde::{Deserialize, Serialize};
use std::{collections::HashMap, path::Path};

/// 困难类型枚举
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize, Copy)]
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

    /// 根据困难类型获取列配置 (身份证列索引, 数据开始行)
    pub fn get_column_config(&self) -> (usize, usize) {
        match self {
            Self::PovertyAlleviatedContinuePolicy | Self::PovertyAlleviatedNoPolicy => (7, 1),
            Self::DisabledWithCertificate => (1, 1),
            Self::RuralMinimumLiving => (6, 2),
            Self::UrbanMinimumLiving => (6, 1),
            Self::RuralSpecialDifficulty => (5, 3),
            Self::AntiPovertyMonitoringRiskNotEliminated
            | Self::AntiPovertyMonitoringRiskEliminated => (11, 1),
            Self::OrphansAndFactuallyUnsupportedChildren => (2, 1),
            Self::LowIncomePopulation => (3, 1),
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

/// 困难人员信息结构
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DifficultPerson {
    pub id_number: String,               // 身份证号
    pub difficulty_type: DifficultyType, // 困难类型
}

/// 匹配结果结构
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MatchResult {
    pub student: Student,
    pub difficult_info: DifficultPerson,
}

/// Excel读取错误类型
#[derive(Debug, Clone, thiserror::Error)]
pub enum ExcelError {
    #[error("File not found: {0}")]
    FileNotFound(String),
    #[error("Read error: {0}")]
    ReadError(String),
    #[error("Parse error: {0}")]
    ParseError(String),
}

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
pub fn read_student_info(file_path: &str) -> Result<Vec<Student>, ExcelError> {
    if !Path::new(file_path).exists() {
        return Err(ExcelError::FileNotFound(file_path.to_string()));
    }

    let mut students = Vec::new();

    if file_path.ends_with(".xls") {
        let mut workbook: Xls<_> =
            open_workbook(file_path).map_err(|e: XlsError| ExcelError::ReadError(e.to_string()))?;
        let range = workbook
            .worksheet_range_at(0)
            .ok_or(ExcelError::ReadError("NO DATA".into()))?
            .map_err(|e| ExcelError::ReadError(e.to_string()))?;

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

                // B列：身份证件号
                let id_value = row.get(1);
                let id_number = id_value
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string();

                if !name.is_empty() && !id_number.is_empty() {
                    let normalized_id = normalize_id_number(&id_number);
                    let student = Student {
                        name: name.clone(),
                        id_number: normalized_id.clone(),
                        // K列：全国学籍号
                        student_id: row
                            .get(10)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // J列：班级
                        class: row
                            .get(9)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // I列：年级
                        grade: row
                            .get(8)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // E列：学校名称
                        school: row
                            .get(4)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                    };
                    students.push(student);
                }
            }
        }
    } else if file_path.ends_with(".xlsx") {
        let mut workbook: Xlsx<_> = open_workbook(file_path)
            .map_err(|e: XlsxError| ExcelError::ReadError(e.to_string()))?;
        let range = workbook
            .worksheet_range_at(0)
            .ok_or(ExcelError::ReadError("NO DATA".into()))?
            .map_err(|e| ExcelError::ReadError(e.to_string()))?;

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

                // B列：身份证件号
                let id_value = row.get(1);
                let id_number = id_value
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string();

                if !name.is_empty() && !id_number.is_empty() {
                    let normalized_id = normalize_id_number(&id_number);
                    let student = Student {
                        name: name.clone(),
                        id_number: normalized_id.clone(),
                        // K列：全国学籍号
                        student_id: row
                            .get(10)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // J列：班级
                        class: row
                            .get(9)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // I列：年级
                        grade: row
                            .get(8)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                        // E列：学校名称
                        school: row
                            .get(4)
                            .and_then(|v| v.as_string())
                            .map(|s| s.trim().to_string()),
                    };
                    students.push(student);
                }
            }
        }
    } else {
        return Err(ExcelError::ReadError("NO DATA".to_string()));
    }
    Ok(students)
}

/// 常规
fn read_common(
    file_path: &str,
    difficulty_type: DifficultyType,
) -> Result<Vec<DifficultPerson>, ExcelError> {
    let mut difficult_people = Vec::new();

    // 根据困难类型确定列位置
    let (id_col, data_start_row) = difficulty_type.get_column_config();

    let mut workbook: Xlsx<_> =
        open_workbook(file_path).map_err(|e: XlsxError| ExcelError::ReadError(e.to_string()))?;
    let range = workbook
        .worksheet_range_at(0)
        .ok_or(ExcelError::ReadError(
            "Cannot find worksheet at index 0".to_string(),
        ))?
        .map_err(|e| ExcelError::ReadError(e.to_string()))?;

    for row in range.rows().skip(data_start_row) {
        let id_number = row
            .get(id_col)
            .and_then(|v| v.as_string())
            .unwrap_or_default()
            .trim()
            .to_string();

        // 只要身份证号不为空就添加记录
        if !id_number.is_empty() {
            let difficult_person = DifficultPerson {
                id_number: normalize_id_number(&id_number),
                difficulty_type,
            };
            difficult_people.push(difficult_person);
        }
    }
    Ok(difficult_people)
}

/// 孤儿
fn read_orphans(file_path: &str) -> Result<Vec<DifficultPerson>, ExcelError> {
    let mut difficult_people = Vec::new();

    let mut workbook: Xls<_> =
        open_workbook(file_path).map_err(|e: XlsError| ExcelError::ReadError(e.to_string()))?;
    let range = workbook
        .worksheet_range_at(0)
        .ok_or(ExcelError::ReadError("NO DATA".to_string()))?
        .map_err(|e| ExcelError::ReadError(e.to_string()))?;

    for row in range.rows().skip(3) {
        let id_number = row
            .get(2)
            .and_then(|v| v.as_string())
            .unwrap_or_default()
            .trim()
            .to_string();

        // 只要身份证号不为空就添加记录
        if !id_number.is_empty() {
            let difficult_person = DifficultPerson {
                id_number: normalize_id_number(&id_number),
                difficulty_type: DifficultyType::OrphansAndFactuallyUnsupportedChildren,
            };
            difficult_people.push(difficult_person);
        }
    }
    let range = workbook
        .worksheet_range_at(2)
        .ok_or(ExcelError::ReadError("NO DATA".to_string()))?
        .map_err(|e| ExcelError::ReadError(e.to_string()))?;

    for row in range.rows().skip(3) {
        let id_number = row
            .get(2)
            .and_then(|v| v.as_string())
            .unwrap_or_default()
            .trim()
            .to_string();

        // 只要身份证号不为空就添加记录
        if !id_number.is_empty() {
            let difficult_person = DifficultPerson {
                id_number: normalize_id_number(&id_number),
                difficulty_type: DifficultyType::OrphansAndFactuallyUnsupportedChildren,
            };
            difficult_people.push(difficult_person);
        }
    }
    Ok(difficult_people)
}

/// 农村低保
fn read_rural_minimum_living(file_path: &str) -> Result<Vec<DifficultPerson>, ExcelError> {
    let mut difficult_people = Vec::new();
    let difficulty_type = DifficultyType::RuralMinimumLiving;

    let mut workbook: Xls<_> = open_workbook(file_path).unwrap();
    let range = workbook
        .worksheet_range_at(1)
        .ok_or(ExcelError::ReadError(
            "Cannot find worksheet at index 0".to_string(),
        ))?
        .map_err(|e| ExcelError::ReadError(e.to_string()))?;
    let id_columns = [6, 15, 17, 19, 21, 23, 25, 27, 29];
    for row in range.rows().skip(2) {
        for col in id_columns {
            let raw_value = row.get(col);
            let id_number = raw_value
                .and_then(|v| v.as_string())
                .unwrap_or_default()
                .trim()
                .to_string();

            if !id_number.is_empty() {
                let normalized_id = normalize_id_number(&id_number);
                let difficult_person = DifficultPerson {
                    id_number: normalized_id,
                    difficulty_type,
                };
                difficult_people.push(difficult_person);
            }
        }
    }
    Ok(difficult_people)
}

/// 城镇低保
fn read_urban_minimum_living(file_path: &str) -> Result<Vec<DifficultPerson>, ExcelError> {
    let mut difficult_people = Vec::new();
    let difficulty_type = DifficultyType::UrbanMinimumLiving;

    let mut workbook: Xls<_> = open_workbook(file_path).unwrap();
    let range = workbook
        .worksheet_range_at(1)
        .ok_or(ExcelError::ReadError(
            "Cannot find worksheet at index 0".to_string(),
        ))?
        .map_err(|e| ExcelError::ReadError(e.to_string()))?;
    let id_columns = [6, 16, 18, 20, 22, 24];
    for row in range.rows().skip(2) {
        for col in id_columns {
            let raw_value = row.get(col);
            let id_number = raw_value
                .and_then(|v| v.as_string())
                .unwrap_or_default()
                .trim()
                .to_string();

            if !id_number.is_empty() {
                let difficult_person = DifficultPerson {
                    id_number: normalize_id_number(&id_number),
                    difficulty_type,
                };
                difficult_people.push(difficult_person);
            }
        }
    }
    Ok(difficult_people)
}

/// 城乡特困
fn read_rural_special_difficulty(file_path: &str) -> Result<Vec<DifficultPerson>, ExcelError> {
    let mut difficult_people = Vec::new();

    let mut workbook: Xlsx<_> =
        open_workbook(file_path).map_err(|e: XlsxError| ExcelError::ReadError(e.to_string()))?;
    let range = workbook
        .worksheet_range_at(1)
        .ok_or(ExcelError::ReadError(
            "Cannot find worksheet at index 1".to_string(),
        ))?
        .map_err(|e| ExcelError::ReadError(e.to_string()))?;

    let id_columns = [5, 26, 31, 33, 35, 37, 39, 41];
    for row in range.rows().skip(3) {
        for col in id_columns {
            let id_number = row
                .get(col)
                .and_then(|v| v.as_string())
                .unwrap_or_default()
                .trim()
                .to_string();

            if !id_number.is_empty() {
                let difficult_person = DifficultPerson {
                    id_number: normalize_id_number(&id_number),
                    difficulty_type: DifficultyType::RuralSpecialDifficulty,
                };
                difficult_people.push(difficult_person);
            }
        }
    }
    Ok(difficult_people)
}

/// 读取困难类型表
pub fn read_difficult_type_table(
    file_path: &str,
    difficulty_type: DifficultyType,
) -> Result<Vec<DifficultPerson>, ExcelError> {
    if !Path::new(file_path).exists() {
        return Err(ExcelError::FileNotFound(file_path.to_string()));
    }

    match difficulty_type {
        DifficultyType::RuralMinimumLiving => read_rural_minimum_living(file_path),
        DifficultyType::RuralSpecialDifficulty => read_rural_special_difficulty(file_path),
        DifficultyType::UrbanMinimumLiving => read_urban_minimum_living(file_path),
        DifficultyType::OrphansAndFactuallyUnsupportedChildren => read_orphans(file_path),
        _ => read_common(file_path, difficulty_type),
    }
}

/// 匹配学生信息和困难类型信息
pub fn match_students_with_difficulty(
    students: &[Student],
    difficult_people: &[DifficultPerson],
) -> Vec<MatchResult> {
    let mut results = Vec::new();

    // 创建学生身份证号的哈希映射以提高查询效率
    let student_map: HashMap<String, &Student> =
        students.iter().map(|s| (s.id_number.clone(), s)).collect();

    for difficult_person in difficult_people {
        if let Some(student) = student_map.get(&difficult_person.id_number) {
            results.push(MatchResult {
                student: (*student).clone(),
                difficult_info: difficult_person.clone(),
            });
        }
    }

    results
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_read() {
        let result = read_rural_minimum_living("/Users/xiaoyu/Downloads/2025年秋季学期困难类型查询专用表及最新的本县户籍特殊困难群体人员信息9.3/3-02025.9.1-2025年9月份农村低保备案表.xls").unwrap();
        println!("农村低保数量 {}", result.len());
        let result = read_rural_special_difficulty("/Users/xiaoyu/Downloads/2025年秋季学期困难类型查询专用表及最新的本县户籍特殊困难群体人员信息9.3/4-2025.9.1-2025年9月份城乡特困人员备案表.xlsx").unwrap();
        println!("特困人员数量 {}", result.len());

        let result = read_urban_minimum_living("/Users/xiaoyu/Downloads/2025年秋季学期困难类型查询专用表及最新的本县户籍特殊困难群体人员信息9.3/5-2025.9.1-2025年9月份城镇低保全部备案表.xls").unwrap();
        println!("城镇低保人员数量 {}", result.len());
    }

    #[test]
    fn test_read_orphans() {
        let file_path = "/Users/xiaoyu/Downloads/2025年秋季学期困难类型查询专用表及最新的本县户籍特殊困难群体人员信息9.3/7-2025.9.1-2025年9月份孤儿及事实无人抚养儿童发放花名册.xls";
        let result = read_orphans(file_path).unwrap();
        println!("数量: {}", result.len());
    }
}
