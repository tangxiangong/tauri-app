# 学生困难类型查找功能使用示例

本文档展示了如何使用新创建的学生困难类型查找功能。

## 功能概述

新增的学生查找功能包括两个主要命令：
- `find_students_by_difficulty`: 根据困难类型查找匹配的学生
- `get_students_match_statistics`: 获取匹配结果的统计信息

## TypeScript 接口

### 导入所需模块

```typescript
import {
  findStudentsByDifficulty,
  getStudentsMatchStatistics,
  executeStudentSearch,
  groupMatchesByDifficultyType,
  searchInMatches,
  exportMatchesToCSV,
  validateSearchParameters,
  type MatchResult,
  type MatchStatistics,
} from '$lib/command';
```

### 基本使用示例

```typescript
async function searchStudents() {
  // 文件路径
  const studentFilePath = '/path/to/student-info.xlsx';
  const difficultyFilePath = '/path/to/difficulty-data.xlsx';
  const difficultyType = '脱贫户(继续享受政策)';

  // 1. 验证参数
  const validation = validateSearchParameters(
    studentFilePath,
    difficultyFilePath,
    difficultyType
  );

  if (!validation.valid) {
    console.error('参数验证失败:', validation.error);
    return;
  }

  // 2. 执行查找
  const result = await findStudentsByDifficulty(
    studentFilePath,
    difficultyFilePath,
    difficultyType
  );

  if (!result.success) {
    console.error('查找失败:', result.error);
    return;
  }

  console.log('找到匹配学生:', result.data?.length || 0);
  console.log('匹配结果:', result.data);
}
```

### 获取统计信息

```typescript
async function getStatistics() {
  const studentFilePath = '/path/to/student-info.xlsx';
  const difficultyFilePath = '/path/to/difficulty-data.xlsx';
  const difficultyType = '持证残疾人';

  const statsResult = await getStudentsMatchStatistics(
    studentFilePath,
    difficultyFilePath,
    difficultyType
  );

  if (statsResult.success && statsResult.data) {
    const stats = statsResult.data;
    console.log(`总学生数: ${stats.total_students}`);
    console.log(`匹配数量: ${stats.total_matches}`);
    console.log('困难类型分布:', stats.difficulty_type_counts);
  }
}
```

### 完整查找示例（推荐）

```typescript
async function completeSearch() {
  const studentFilePath = '/path/to/student-info.xlsx';
  const difficultyFilePath = '/path/to/difficulty-data.xlsx';
  const difficultyType = '农村低保';

  // 使用 executeStudentSearch 一次性获取所有信息
  const result = await executeStudentSearch(
    studentFilePath,
    difficultyFilePath,
    difficultyType
  );

  if (!result.success) {
    console.error('查找失败:', result.error);
    return;
  }

  console.log('=== 查找结果 ===');
  console.log(`找到 ${result.matches.length} 名匹配学生`);
  
  console.log('=== 统计信息 ===');
  console.log(`总学生数: ${result.statistics.total_students}`);
  console.log(`匹配数量: ${result.statistics.total_matches}`);
  
  // 按困难类型分组
  const groupedByType = groupMatchesByDifficultyType(result.matches);
  console.log('=== 按困难类型分组 ===');
  Object.entries(groupedByType).forEach(([type, matches]) => {
    console.log(`${type}: ${matches.length} 人`);
  });
}
```

### 搜索和过滤功能

```typescript
async function searchAndFilter() {
  // 首先获取所有匹配结果
  const result = await findStudentsByDifficulty(
    '/path/to/student-info.xlsx',
    '/path/to/difficulty-data.xlsx',
    '城镇低保'
  );

  if (result.success && result.data) {
    const allMatches = result.data;

    // 搜索特定学生
    const searchResults = searchInMatches(allMatches, '张');
    console.log(`姓氏包含"张"的学生: ${searchResults.length} 人`);

    // 按学校分组
    const groupedBySchool = groupMatchesBySchool(allMatches);
    console.log('=== 按学校分组 ===');
    Object.entries(groupedBySchool).forEach(([school, matches]) => {
      console.log(`${school}: ${matches.length} 人`);
    });

    // 按年级分组
    const groupedByGrade = groupMatchesByGrade(allMatches);
    console.log('=== 按年级分组 ===');
    Object.entries(groupedByGrade).forEach(([grade, matches]) => {
      console.log(`${grade}: ${matches.length} 人`);
    });
  }
}
```

### 导出数据

```typescript
async function exportData() {
  const result = await findStudentsByDifficulty(
    '/path/to/student-info.xlsx',
    '/path/to/difficulty-data.xlsx',
    '孤儿及事实无人抚养儿童'
  );

  if (result.success && result.data) {
    // 导出为 CSV
    const csvData = exportMatchesToCSV(result.data);
    
    // 保存到文件或下载
    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', '匹配结果.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
}
```

## Svelte 组件示例

```svelte
<script lang="ts">
  import {
    executeStudentSearch,
    openFileDialog,
    getDifficultyTypeOptions,
    type MatchResult,
    type MatchStatistics,
    type DifficultyTypeOption
  } from '$lib/command';

  let studentFilePath = '';
  let difficultyFilePath = '';
  let selectedDifficultyType = '';
  let difficultyOptions: DifficultyTypeOption[] = [];
  let matches: MatchResult[] = [];
  let statistics: MatchStatistics | null = null;
  let loading = false;
  let error = '';

  // 加载困难类型选项
  async function loadDifficultyOptions() {
    const result = await getDifficultyTypeOptions();
    if (result.success && result.data) {
      difficultyOptions = result.data;
    }
  }

  // 选择学生文件
  async function selectStudentFile() {
    const path = await openFileDialog('选择学生信息表文件');
    if (path) {
      studentFilePath = path;
    }
  }

  // 选择困难类型文件
  async function selectDifficultyFile() {
    const path = await openFileDialog('选择困难类型数据表文件');
    if (path) {
      difficultyFilePath = path;
    }
  }

  // 执行查找
  async function handleSearch() {
    if (!studentFilePath || !difficultyFilePath || !selectedDifficultyType) {
      error = '请选择所有必需的文件和困难类型';
      return;
    }

    loading = true;
    error = '';

    try {
      const result = await executeStudentSearch(
        studentFilePath,
        difficultyFilePath,
        selectedDifficultyType
      );

      if (result.success) {
        matches = result.matches;
        statistics = result.statistics;
      } else {
        error = result.error || '查找失败';
      }
    } catch (err) {
      error = `查找时发生错误: ${err}`;
    } finally {
      loading = false;
    }
  }

  // 页面加载时获取困难类型选项
  loadDifficultyOptions();
</script>

<div class="container">
  <h1>学生困难类型查找</h1>

  <div class="form-section">
    <div class="form-group">
      <label>学生信息表文件:</label>
      <button on:click={selectStudentFile}>选择文件</button>
      {#if studentFilePath}
        <span class="file-path">{studentFilePath}</span>
      {/if}
    </div>

    <div class="form-group">
      <label>困难类型数据表文件:</label>
      <button on:click={selectDifficultyFile}>选择文件</button>
      {#if difficultyFilePath}
        <span class="file-path">{difficultyFilePath}</span>
      {/if}
    </div>

    <div class="form-group">
      <label for="difficulty-type">困难类型:</label>
      <select id="difficulty-type" bind:value={selectedDifficultyType}>
        <option value="">请选择困难类型</option>
        {#each difficultyOptions as option}
          <option value={option.value}>{option.label}</option>
        {/each}
      </select>
    </div>

    <button on:click={handleSearch} disabled={loading}>
      {loading ? '查找中...' : '开始查找'}
    </button>
  </div>

  {#if error}
    <div class="error">{error}</div>
  {/if}

  {#if statistics}
    <div class="statistics">
      <h2>统计信息</h2>
      <p>总学生数: {statistics.total_students}</p>
      <p>匹配数量: {statistics.total_matches}</p>
      
      <h3>困难类型分布:</h3>
      <ul>
        {#each Object.entries(statistics.difficulty_type_counts) as [type, count]}
          <li>{type}: {count} 人</li>
        {/each}
      </ul>
    </div>
  {/if}

  {#if matches.length > 0}
    <div class="results">
      <h2>匹配结果</h2>
      <table>
        <thead>
          <tr>
            <th>姓名</th>
            <th>身份证号</th>
            <th>学号</th>
            <th>班级</th>
            <th>年级</th>
            <th>学校</th>
            <th>困难类型</th>
          </tr>
        </thead>
        <tbody>
          {#each matches as match}
            <tr>
              <td>{match.student.name}</td>
              <td>{match.student.id_number}</td>
              <td>{match.student.student_id || '-'}</td>
              <td>{match.student.class || '-'}</td>
              <td>{match.student.grade || '-'}</td>
              <td>{match.student.school || '-'}</td>
              <td>{match.difficult_info.difficulty_type}</td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
  {/if}
</div>

<style>
  .container {
    max-width: 1200px;
    margin: 0 auto;
    padding: 20px;
  }

  .form-section {
    margin-bottom: 30px;
  }

  .form-group {
    margin-bottom: 15px;
  }

  .form-group label {
    display: block;
    margin-bottom: 5px;
    font-weight: bold;
  }

  .form-group button,
  .form-group select {
    padding: 8px 12px;
    border: 1px solid #ddd;
    border-radius: 4px;
  }

  .file-path {
    margin-left: 10px;
    font-size: 0.9em;
    color: #666;
  }

  .error {
    color: red;
    background-color: #ffebee;
    padding: 10px;
    border-radius: 4px;
    margin-bottom: 20px;
  }

  .statistics {
    background-color: #f5f5f5;
    padding: 15px;
    border-radius: 4px;
    margin-bottom: 20px;
  }

  .results table {
    width: 100%;
    border-collapse: collapse;
  }

  .results th,
  .results td {
    padding: 8px 12px;
    text-align: left;
    border-bottom: 1px solid #ddd;
  }

  .results th {
    background-color: #f5f5f5;
    font-weight: bold;
  }

  button:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }
</style>
```

## 注意事项

1. **文件路径**: 确保提供的文件路径是绝对路径且文件存在
2. **困难类型**: 困难类型必须与系统定义的枚举值完全匹配
3. **Excel 格式**: 支持 .xlsx 和 .xls 格式的文件
4. **错误处理**: 始终检查命令返回的 `success` 字段和 `error` 信息
5. **性能**: 对于大型文件，查找操作可能需要一些时间，建议显示加载状态

## 可用的困难类型

- 脱贫户(继续享受政策)
- 脱贫户(不享受政策)
- 持证残疾人
- 农村低保
- 城镇低保
- 城乡特困
- 防返贫监测对象(风险未消除)
- 防返贫监测对象(风险已消除)
- 孤儿及事实无人抚养儿童
- 低收入人口