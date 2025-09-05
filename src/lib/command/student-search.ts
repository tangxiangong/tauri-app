import { invoke } from "@tauri-apps/api/core";
import type {
  CommandResult,
  DifficultStudent,
  DifficultyType,
  MatchResult,
  Student,
} from "./upload";

/**
 * 匹配结果统计信息
 */
export interface MatchStatistics {
  total_students: number;
  total_matches: number;
  difficulty_type_counts: Partial<Record<DifficultyType, number>>;
}

/**
 * 根据困难类型查找学生信息
 * @param studentFilePath 学生信息表文件路径
 * @param difficultyFilePath 困难类型数据表文件路径
 * @param difficultyType 困难类型
 * @returns 匹配的学生结果列表
 */
export async function findStudentsByDifficulty(
  studentFilePath: string,
  difficultyFilePath: string,
  difficultyType: string,
): Promise<CommandResult<MatchResult[]>> {
  return await invoke("find_students_by_difficulty", {
    studentFilePath,
    difficultyFilePath,
    difficultyType,
  });
}

/**
 * 获取学生匹配统计信息
 * @param studentFilePath 学生信息表文件路径
 * @param difficultyFilePath 困难类型数据表文件路径
 * @param difficultyType 困难类型
 * @returns 匹配统计信息
 */
export async function getStudentsMatchStatistics(
  studentFilePath: string,
  difficultyFilePath: string,
  difficultyType: string,
): Promise<CommandResult<MatchStatistics>> {
  return await invoke("get_students_match_statistics", {
    studentFilePath,
    difficultyFilePath,
    difficultyType,
  });
}

/**
 * 执行学生查找并获取详细结果
 * @param studentFilePath 学生信息表文件路径
 * @param difficultyFilePath 困难类型数据表文件路径
 * @param difficultyType 困难类型
 * @returns 包含匹配结果和统计信息的完整数据
 */
export async function executeStudentSearch(
  studentFilePath: string,
  difficultyFilePath: string,
  difficultyType: string,
): Promise<{
  matches: MatchResult[];
  statistics: MatchStatistics;
  success: boolean;
  error?: string;
}> {
  try {
    // 并行执行查找和统计
    const [matchResult, statsResult] = await Promise.all([
      findStudentsByDifficulty(
        studentFilePath,
        difficultyFilePath,
        difficultyType,
      ),
      getStudentsMatchStatistics(
        studentFilePath,
        difficultyFilePath,
        difficultyType,
      ),
    ]);

    if (!matchResult.success) {
      return {
        matches: [],
        statistics: {
          total_students: 0,
          total_matches: 0,
          difficulty_type_counts: {} as Partial<Record<DifficultyType, number>>,
        },
        success: false,
        error: matchResult.error || "查找学生失败",
      };
    }

    if (!statsResult.success) {
      return {
        matches: matchResult.data || [],
        statistics: {
          total_students: 0,
          total_matches: 0,
          difficulty_type_counts: {} as Partial<Record<DifficultyType, number>>,
        },
        success: false,
        error: statsResult.error || "获取统计信息失败",
      };
    }

    return {
      matches: matchResult.data || [],
      statistics: statsResult.data ||
        {
          total_students: 0,
          total_matches: 0,
          difficulty_type_counts: {} as Partial<Record<DifficultyType, number>>,
        },
      success: true,
    };
  } catch (error) {
    return {
      matches: [],
      statistics: {
        total_students: 0,
        total_matches: 0,
        difficulty_type_counts: {} as Partial<Record<DifficultyType, number>>,
      },
      success: false,
      error: `执行查找时发生错误: ${error}`,
    };
  }
}

/**
 * 按困难类型分组匹配结果
 * @param matches 匹配结果数组
 * @returns 按困难类型分组的结果
 */
export function groupMatchesByDifficultyType(
  matches: MatchResult[],
): Record<DifficultyType, MatchResult[]> {
  const grouped: Record<string, MatchResult[]> = {};

  matches.forEach((match) => {
    const type = match.difficult_info.difficulty_type;
    if (!grouped[type]) {
      grouped[type] = [];
    }
    grouped[type].push(match);
  });

  return grouped as Record<DifficultyType, MatchResult[]>;
}

/**
 * 按学校分组匹配结果
 * @param matches 匹配结果数组
 * @returns 按学校分组的结果
 */
export function groupMatchesBySchool(
  matches: MatchResult[],
): Record<string, MatchResult[]> {
  const grouped: Record<string, MatchResult[]> = {};

  matches.forEach((match) => {
    const school = match.student.school || "未知学校";
    if (!grouped[school]) {
      grouped[school] = [];
    }
    grouped[school].push(match);
  });

  return grouped;
}

/**
 * 按年级分组匹配结果
 * @param matches 匹配结果数组
 * @returns 按年级分组的结果
 */
export function groupMatchesByGrade(
  matches: MatchResult[],
): Record<string, MatchResult[]> {
  const grouped: Record<string, MatchResult[]> = {};

  matches.forEach((match) => {
    const grade = match.student.grade || "未知年级";
    if (!grouped[grade]) {
      grouped[grade] = [];
    }
    grouped[grade].push(match);
  });

  return grouped;
}

/**
 * 搜索匹配结果中的特定学生
 * @param matches 匹配结果数组
 * @param searchTerm 搜索关键词（可以是姓名或身份证号的一部分）
 * @returns 匹配的结果
 */
export function searchInMatches(
  matches: MatchResult[],
  searchTerm: string,
): MatchResult[] {
  if (!searchTerm.trim()) {
    return matches;
  }

  const term = searchTerm.toLowerCase().trim();
  return matches.filter((match) => {
    return (
      match.student.name.toLowerCase().includes(term) ||
      match.student.id_number.includes(term) ||
      (match.student.student_id &&
        match.student.student_id.toLowerCase().includes(term)) ||
      (match.student.class && match.student.class.toLowerCase().includes(term))
    );
  });
}

/**
 * 验证查找参数
 * @param studentFilePath 学生文件路径
 * @param difficultyFilePath 困难类型文件路径
 * @param difficultyType 困难类型
 * @returns 验证结果
 */
export function validateSearchParameters(
  studentFilePath: string,
  difficultyFilePath: string,
  difficultyType: string,
): { valid: boolean; error?: string } {
  if (!studentFilePath.trim()) {
    return { valid: false, error: "请选择学生信息表文件" };
  }

  if (!difficultyFilePath.trim()) {
    return { valid: false, error: "请选择困难类型数据表文件" };
  }

  if (!difficultyType.trim()) {
    return { valid: false, error: "请选择困难类型" };
  }

  return { valid: true };
}

/**
 * 导出匹配结果到 Excel 文件
 * @param matches 匹配结果数组
 * @param outputPath 输出文件路径
 * @returns 导出结果
 */
export async function exportMatchesToExcel(
  matches: MatchResult[],
  outputPath: string,
): Promise<CommandResult<string>> {
  return await invoke("export_matches_to_excel", {
    matches,
    outputPath,
  });
}
