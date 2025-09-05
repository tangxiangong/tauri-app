import { invoke } from "@tauri-apps/api/core";
import type { CommandResult, DifficultyType, MatchResult } from "./upload.ts";

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
