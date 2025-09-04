import { invoke } from "@tauri-apps/api/core";
import { open } from "@tauri-apps/plugin-dialog";

// 类型定义
export interface CommandResult<T> {
  success: boolean;
  data: T | null;
  error: string | null;
}

export interface FileInfo {
  name: string;
  path: string;
  size: number;
  extension: string;
}

export interface DifficultyTypeOption {
  label: string;
  value: string;
}

export interface Student {
  name: string;
  id_number: string;
  student_id?: string;
  class?: string;
  grade?: string;
  school?: string;
}

export interface DifficultStudent {
  name: string;
  id_number: string;
  difficulty_type: string;
  source_file: string;
  extra_info: Record<string, string>;
}

export interface MatchResult {
  student: Student;
  difficult_info: DifficultStudent;
}

/**
 * 处理上传的学生文件和困难类型文件
 * @param studentFilePath 学生文件路径
 * @param difficultyFilePath 困难类型文件路径
 * @param difficultyType 困难类型
 * @returns 匹配结果
 */
export async function processUploadedFiles(
  studentFilePath: string,
  difficultyFilePath: string,
  difficultyType: string,
): Promise<CommandResult<MatchResult[]>> {
  return await invoke("process_uploaded_files", {
    studentFilePath,
    difficultyFilePath,
    difficultyType,
  });
}

/**
 * 验证上传的文件
 * @param filePath 文件路径
 * @returns 文件信息
 */
export async function validateUploadedFile(
  filePath: string,
): Promise<CommandResult<FileInfo>> {
  return await invoke("validate_uploaded_file", {
    filePath,
  });
}

/**
 * 获取困难类型选项列表
 * @returns 困难类型选项列表
 */
export async function getDifficultyTypeOptions(): Promise<
  CommandResult<DifficultyTypeOption[]>
> {
  return await invoke("get_difficulty_type_options");
}

/**
 * 打开文件选择对话框
 * @param title 对话框标题
 * @param extensions 允许的文件扩展名
 * @returns 选择的文件路径
 */
export async function openFileDialog(
  title: string,
  extensions: string[] = ["xlsx", "xls"],
): Promise<string | null> {
  try {
    const selected = await open({
      title,
      multiple: false,
      filters: [
        {
          name: "Excel 文件",
          extensions,
        },
      ],
    });

    return Array.isArray(selected) ? selected[0] : selected;
  } catch (error) {
    return null;
  }
}

/**
 * 格式化文件大小
 * @param bytes 字节数
 * @returns 格式化后的文件大小字符串
 */
export function formatFileSize(bytes: number): string {
  if (bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

/**
 * 验证文件是否为 Excel 文件
 * @param fileName 文件名
 * @returns 是否为 Excel 文件
 */
export function isExcelFile(fileName: string): boolean {
  const extension = fileName.toLowerCase().split(".").pop();
  return extension === "xlsx" || extension === "xls";
}

/**
 * 生成匹配结果的摘要信息
 * @param results 匹配结果数组
 * @returns 摘要信息
 */
export function generateMatchSummary(results: MatchResult[]): {
  totalMatches: number;
  difficultyTypeCounts: Record<string, number>;
} {
  const difficultyTypeCounts: Record<string, number> = {};

  results.forEach((result) => {
    const type = result.difficult_info.difficulty_type;
    difficultyTypeCounts[type] = (difficultyTypeCounts[type] || 0) + 1;
  });

  return {
    totalMatches: results.length,
    difficultyTypeCounts,
  };
}

/**
 * 掩码身份证号以保护隐私
 * @param idNumber 身份证号
 * @returns 掩码后的身份证号
 */
export function maskIdNumber(idNumber: string): string {
  if (idNumber.length >= 6) {
    return `${idNumber.substring(0, 3)}****${
      idNumber.substring(idNumber.length - 3)
    }`;
  }
  return "****";
}
