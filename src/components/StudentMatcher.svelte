<script lang="ts">
    import { onMount } from "svelte";
    import {
        validateUploadedFile,
        getDifficultyTypeOptions,
        openFileDialog,
        saveFileDialog,
        formatFileSize,
        maskIdNumber,
        generateMatchSummary,
        type FileInfo,
        type DifficultyTypeOption,
        type MatchResult,
    } from "$lib/command/upload";
    import {
        executeStudentSearch,
        exportMatchesToExcel,
        type MatchStatistics,
    } from "$lib/command/student-search";

    // 响应式状态
    let studentFile: FileInfo | null = $state(null);
    let difficultyFile: FileInfo | null = $state(null);
    let selectedDifficultyType: string = $state("");
    let difficultyTypeOptions: DifficultyTypeOption[] = $state([]);
    let matchResults: MatchResult[] = $state([]);
    let matchStatistics: MatchStatistics | null = $state(null);
    let isLoading: boolean = $state(false);
    let error: string = $state("");
    let success: string = $state("");
    let hasSearched: boolean = $state(false);
    let isExporting: boolean = $state(false);

    // 分页状态
    let currentPage: number = $state(1);
    let itemsPerPage: number = $state(10);

    // 计算属性
    let canMatch = $derived(
        studentFile && difficultyFile && selectedDifficultyType,
    );
    let matchSummary = $derived(generateMatchSummary(matchResults));
    let totalPages = $derived(Math.ceil(matchResults.length / itemsPerPage));
    let paginatedResults = $derived.by(() => {
        const start = (currentPage - 1) * itemsPerPage;
        const end = start + itemsPerPage;
        return matchResults.slice(start, end);
    });

    // 初始化
    onMount(async () => {
        await loadDifficultyTypeOptions();
    });

    // 加载困难类型选项
    async function loadDifficultyTypeOptions() {
        try {
            const result = await getDifficultyTypeOptions();
            if (result.success && result.data) {
                difficultyTypeOptions = result.data;
            } else {
                setError(
                    "获取困难类型选项失败: " + (result.error || "未知错误"),
                );
            }
        } catch (err) {
            setError("获取困难类型选项失败: " + String(err));
        }
    }

    // 选择学生文件
    async function selectStudentFile() {
        try {
            const filePath = await openFileDialog("选择学生信息表");

            if (filePath) {
                await validateAndSetFile(filePath, "student");
            }
        } catch (err) {
            setError("选择学生文件失败: " + String(err));
        }
    }

    // 选择困难类型文件
    async function selectDifficultyFile() {
        try {
            const filePath = await openFileDialog("选择困难类型表");

            if (filePath) {
                await validateAndSetFile(filePath, "difficulty");
            }
        } catch (err) {
            setError("选择困难类型文件失败: " + String(err));
        }
    }

    // 验证并设置文件
    async function validateAndSetFile(
        filePath: string,
        fileType: "student" | "difficulty",
    ) {
        try {
            const result = await validateUploadedFile(filePath);
            if (result.success && result.data) {
                if (fileType === "student") {
                    studentFile = result.data;
                    setSuccess("学生文件验证成功");
                } else {
                    difficultyFile = result.data;
                    setSuccess("困难类型文件验证成功");
                }
            } else {
                setError("文件验证失败: " + (result.error || "未知错误"));
            }
        } catch (err) {
            setError("文件验证失败: " + String(err));
        }
    }

    // 执行匹配
    async function performMatch() {
        if (!canMatch) return;

        isLoading = true;
        error = "";
        success = "";

        try {
            const result = await executeStudentSearch(
                studentFile!.path,
                difficultyFile!.path,
                selectedDifficultyType,
            );

            if (result.success) {
                matchResults = result.matches;
                matchStatistics = result.statistics;
                hasSearched = true;
                resetPagination();
                setSuccess(
                    `匹配完成！找到 ${result.matches.length} 个匹配结果`,
                );
            } else {
                setError("匹配失败: " + (result.error || "未知错误"));
                hasSearched = false;
            }
        } catch (err) {
            setError("匹配失败: " + String(err));
            hasSearched = false;
        } finally {
            isLoading = false;
        }
    }

    // 重置所有数据
    function resetAll() {
        studentFile = null;
        difficultyFile = null;
        selectedDifficultyType = "";
        matchResults = [];
        matchStatistics = null;
        hasSearched = false;
        error = "";
        success = "";
        isExporting = false;
        resetPagination();
    }

    // 分页控制函数
    function goToPage(page: number) {
        if (page >= 1 && page <= totalPages) {
            currentPage = page;
        }
    }

    function nextPage() {
        if (currentPage < totalPages) {
            currentPage++;
        }
    }

    function prevPage() {
        if (currentPage > 1) {
            currentPage--;
        }
    }

    function resetPagination() {
        currentPage = 1;
    }

    // 设置错误信息
    function setError(message: string) {
        error = message;
        success = "";
        setTimeout(() => (error = ""), 5000);
    }

    // 设置成功信息
    function setSuccess(message: string) {
        success = message;
        error = "";
        setTimeout(() => (success = ""), 3000);
    }

    // 导出到 Excel
    async function exportToExcel() {
        if (matchResults.length === 0) {
            setError("没有可导出的数据");
            return;
        }

        try {
            isExporting = true;

            // 生成文件名
            const timestamp = new Date()
                .toISOString()
                .slice(0, 19)
                .replace(/:/g, "-");
            const filename = `学生困难类型匹配结果_${selectedDifficultyType}_${timestamp}.xlsx`;

            // 使用保存对话框让用户选择保存位置
            const savePath = await saveFileDialog("保存 Excel 文件", filename, [
                "xlsx",
            ]);

            if (!savePath) {
                isExporting = false;
                return;
            }

            const outputPath = savePath.endsWith(".xlsx")
                ? savePath
                : `${savePath}.xlsx`;

            const result = await exportMatchesToExcel(matchResults, outputPath);

            if (result.success) {
                setSuccess(
                    `成功导出 ${matchResults.length} 条记录到 Excel 文件`,
                );
            } else {
                setError("导出失败: " + (result.error || "未知错误"));
            }
        } catch (err) {
            setError("导出失败: " + String(err));
        } finally {
            isExporting = false;
        }
    }
</script>

<div class="container mx-auto p-6 max-w-6xl">
    <div class="card bg-base-100 shadow-xl">
        <div class="card-body">
            <h1 class="card-title text-2xl mb-6 justify-center">
                学生困难类型匹配
            </h1>

            <!-- 错误和成功提示 -->
            {#if error}
                <div class="alert alert-error mb-4">
                    <svg
                        xmlns="http://www.w3.org/2000/svg"
                        class="stroke-current shrink-0 h-6 w-6"
                        fill="none"
                        viewBox="0 0 24 24"
                    >
                        <path
                            stroke-linecap="round"
                            stroke-linejoin="round"
                            stroke-width="2"
                            d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z"
                        />
                    </svg>
                    <span>{error}</span>
                </div>
            {/if}

            {#if success}
                <div class="alert alert-success mb-4">
                    <svg
                        xmlns="http://www.w3.org/2000/svg"
                        class="stroke-current shrink-0 h-6 w-6"
                        fill="none"
                        viewBox="0 0 24 24"
                    >
                        <path
                            stroke-linecap="round"
                            stroke-linejoin="round"
                            stroke-width="2"
                            d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                        />
                    </svg>
                    <span>{success}</span>
                </div>
            {/if}

            <!-- 文件上传区域 -->
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
                <!-- 学生文件上传 -->
                <div class="card bg-base-200">
                    <div class="card-body">
                        <h2 class="card-title text-lg">
                            <svg
                                xmlns="http://www.w3.org/2000/svg"
                                class="h-5 w-5"
                                fill="none"
                                viewBox="0 0 24 24"
                                stroke="currentColor"
                            >
                                <path
                                    stroke-linecap="round"
                                    stroke-linejoin="round"
                                    stroke-width="2"
                                    d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197m13.5-9a2.5 2.5 0 11-5 0 2.5 2.5 0 015 0z"
                                />
                            </svg>
                            学生信息表
                        </h2>
                        {#if studentFile}
                            <div class="bg-success/20 p-3 rounded-lg">
                                <p class="font-semibold text-success">
                                    {studentFile.name}
                                </p>
                                <p class="text-sm opacity-70">
                                    大小: {formatFileSize(studentFile.size)}
                                </p>
                            </div>
                        {:else}
                            <div
                                class="text-center py-8 border-2 border-dashed border-base-300 rounded-lg"
                            >
                                <p class="text-base-content/70 mb-4">
                                    选择包含学生信息的 Excel 文件
                                </p>
                            </div>
                        {/if}
                        <button
                            class="btn btn-primary"
                            onclick={selectStudentFile}
                            disabled={isLoading}
                        >
                            <svg
                                xmlns="http://www.w3.org/2000/svg"
                                class="h-4 w-4"
                                fill="none"
                                viewBox="0 0 24 24"
                                stroke="currentColor"
                            >
                                <path
                                    stroke-linecap="round"
                                    stroke-linejoin="round"
                                    stroke-width="2"
                                    d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                                />
                            </svg>
                            {studentFile ? "重新选择" : "选择文件"}
                        </button>
                    </div>
                </div>

                <!-- 困难类型文件上传 -->
                <div class="card bg-base-200">
                    <div class="card-body">
                        <h2 class="card-title text-lg">
                            <svg
                                xmlns="http://www.w3.org/2000/svg"
                                class="h-5 w-5"
                                fill="none"
                                viewBox="0 0 24 24"
                                stroke="currentColor"
                            >
                                <path
                                    stroke-linecap="round"
                                    stroke-linejoin="round"
                                    stroke-width="2"
                                    d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                                />
                            </svg>
                            困难类型表
                        </h2>
                        {#if difficultyFile}
                            <div class="bg-success/20 p-3 rounded-lg">
                                <p class="font-semibold text-success">
                                    {difficultyFile.name}
                                </p>
                                <p class="text-sm opacity-70">
                                    大小: {formatFileSize(difficultyFile.size)}
                                </p>
                            </div>
                        {:else}
                            <div
                                class="text-center py-8 border-2 border-dashed border-base-300 rounded-lg"
                            >
                                <p class="text-base-content/70 mb-4">
                                    选择困难类型信息的 Excel 文件
                                </p>
                            </div>
                        {/if}
                        <button
                            class="btn btn-primary"
                            onclick={selectDifficultyFile}
                            disabled={isLoading}
                        >
                            <svg
                                xmlns="http://www.w3.org/2000/svg"
                                class="h-4 w-4"
                                fill="none"
                                viewBox="0 0 24 24"
                                stroke="currentColor"
                            >
                                <path
                                    stroke-linecap="round"
                                    stroke-linejoin="round"
                                    stroke-width="2"
                                    d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                                />
                            </svg>
                            {difficultyFile ? "重新选择" : "选择文件"}
                        </button>
                    </div>
                </div>
            </div>

            <!-- 困难类型选择 -->
            <div class="mb-6">
                <label class="form-control w-full max-w-xs">
                    <div class="label">
                        <span class="label-text font-semibold"
                            >选择困难类型</span
                        >
                    </div>
                    <select
                        class="select select-bordered"
                        bind:value={selectedDifficultyType}
                        disabled={isLoading}
                    >
                        <option value="">请选择困难类型</option>
                        {#each difficultyTypeOptions as option}
                            <option value={option.value}>{option.label}</option>
                        {/each}
                    </select>
                </label>
            </div>

            <!-- 操作按钮 -->
            <div class="flex gap-4 mb-6">
                <button
                    class="btn btn-success btn-lg"
                    class:btn-disabled={!canMatch || isLoading}
                    onclick={performMatch}
                >
                    {#if isLoading}
                        <span class="loading loading-spinner loading-sm"></span>
                        匹配中...
                    {:else}
                        <svg
                            xmlns="http://www.w3.org/2000/svg"
                            class="h-5 w-5"
                            fill="none"
                            viewBox="0 0 24 24"
                            stroke="currentColor"
                        >
                            <path
                                stroke-linecap="round"
                                stroke-linejoin="round"
                                stroke-width="2"
                                d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"
                            />
                        </svg>
                        开始匹配
                    {/if}
                </button>

                <button class="btn btn-secondary" onclick={resetAll}>
                    <svg
                        xmlns="http://www.w3.org/2000/svg"
                        class="h-4 w-4"
                        fill="none"
                        viewBox="0 0 24 24"
                        stroke="currentColor"
                    >
                        <path
                            stroke-linecap="round"
                            stroke-linejoin="round"
                            stroke-width="2"
                            d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"
                        />
                    </svg>
                    重置
                </button>

                {#if hasSearched}
                    <button
                        class="btn btn-accent"
                        onclick={exportToExcel}
                        disabled={isExporting}
                    >
                        {#if isExporting}
                            <span class="loading loading-spinner loading-sm"
                            ></span>
                            导出中...
                        {:else}
                            <svg
                                xmlns="http://www.w3.org/2000/svg"
                                class="h-4 w-4"
                                fill="none"
                                viewBox="0 0 24 24"
                                stroke="currentColor"
                            >
                                <path
                                    stroke-linecap="round"
                                    stroke-linejoin="round"
                                    stroke-width="2"
                                    d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                                />
                            </svg>
                            导出 Excel
                        {/if}
                    </button>
                {/if}
            </div>

            <!-- 匹配结果统计 -->
            {#if hasSearched && matchStatistics}
                <div class="stats shadow mb-6">
                    <div class="stat">
                        <div class="stat-title">匹配数量</div>
                        <div class="stat-value text-primary">
                            {matchStatistics.total_matches}
                        </div>
                    </div>
                    <div class="stat">
                        <div class="stat-title">困难类型</div>
                        <div class="stat-value text-secondary text-sm">
                            {selectedDifficultyType}
                        </div>
                    </div>
                </div>

                {#if matchResults.length === 0}
                    <!-- 无匹配结果提示 -->
                    <div class="card bg-base-200">
                        <div class="card-body text-center py-12">
                            <svg
                                xmlns="http://www.w3.org/2000/svg"
                                class="h-16 w-16 mx-auto text-base-content/30 mb-4"
                                fill="none"
                                viewBox="0 0 24 24"
                                stroke="currentColor"
                            >
                                <path
                                    stroke-linecap="round"
                                    stroke-linejoin="round"
                                    stroke-width="2"
                                    d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                                />
                            </svg>
                            <h3
                                class="text-xl font-semibold text-base-content/70 mb-2"
                            >
                                未找到匹配结果
                            </h3>
                            <p class="text-base-content/50">
                                没有找到符合 "{selectedDifficultyType}"
                                困难类型的学生
                            </p>
                            <p class="text-base-content/40 text-sm mt-2">
                                请检查困难类型数据表是否包含相关学生信息，或尝试其他困难类型
                            </p>
                        </div>
                    </div>
                {:else}
                    <!-- 匹配结果列表 -->
                    <div class="card bg-base-200">
                        <div class="card-body">
                            <h3 class="card-title text-xl mb-4">
                                <svg
                                    xmlns="http://www.w3.org/2000/svg"
                                    class="h-6 w-6"
                                    fill="none"
                                    viewBox="0 0 24 24"
                                    stroke="currentColor"
                                >
                                    <path
                                        stroke-linecap="round"
                                        stroke-linejoin="round"
                                        stroke-width="2"
                                        d="M9 5H7a2 2 0 00-2 2v10a2 2 0 002 2h8a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"
                                    />
                                </svg>
                                匹配结果 ({matchResults.length} 人)
                            </h3>

                            <div class="overflow-x-auto">
                                <table class="table table-zebra">
                                    <thead>
                                        <tr>
                                            <th>序号</th>
                                            <th>学生姓名</th>
                                            <th>身份证号</th>
                                            <th>班级</th>
                                            <th>年级</th>
                                            <th>学校</th>
                                            <th>困难类型</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {#each paginatedResults as result, index}
                                            <tr class="hover">
                                                <td
                                                    >{(currentPage - 1) *
                                                        itemsPerPage +
                                                        index +
                                                        1}</td
                                                >
                                                <td class="font-semibold"
                                                    >{result.student.name}</td
                                                >
                                                <td class="font-mono text-sm"
                                                    >{maskIdNumber(
                                                        result.student
                                                            .id_number,
                                                    )}</td
                                                >
                                                <td
                                                    >{result.student.class ||
                                                        "未知"}</td
                                                >
                                                <td
                                                    >{result.student.grade ||
                                                        "未知"}</td
                                                >
                                                <td
                                                    >{result.student.school ||
                                                        "未知"}</td
                                                >
                                                <td>
                                                    <div
                                                        class="badge badge-primary badge-sm"
                                                    >
                                                        {result.difficult_info
                                                            .difficulty_type}
                                                    </div>
                                                </td>
                                            </tr>
                                        {/each}
                                    </tbody>
                                </table>
                            </div>

                            <!-- 分页控件 -->
                            {#if totalPages > 1}
                                <div class="flex justify-center mt-6">
                                    <div class="join">
                                        <button
                                            class="join-item btn"
                                            onclick={prevPage}
                                            disabled={currentPage === 1}
                                        >
                                            «
                                        </button>

                                        {#each Array.from({ length: totalPages }, (_, i) => i + 1) as page}
                                            {#if totalPages <= 10 || page <= 3 || page >= totalPages - 2 || Math.abs(page - currentPage) <= 1}
                                                <button
                                                    class="join-item btn {currentPage ===
                                                    page
                                                        ? 'btn-active'
                                                        : ''}"
                                                    onclick={() =>
                                                        goToPage(page)}
                                                >
                                                    {page}
                                                </button>
                                            {:else if page === 4 && currentPage > 5}
                                                <button
                                                    class="join-item btn btn-disabled"
                                                    >...</button
                                                >
                                            {:else if page === totalPages - 3 && currentPage < totalPages - 4}
                                                <button
                                                    class="join-item btn btn-disabled"
                                                    >...</button
                                                >
                                            {/if}
                                        {/each}

                                        <button
                                            class="join-item btn"
                                            onclick={nextPage}
                                            disabled={currentPage ===
                                                totalPages}
                                        >
                                            »
                                        </button>
                                    </div>
                                </div>

                                <div
                                    class="text-center mt-4 text-sm opacity-70"
                                >
                                    第 {currentPage} 页，共 {totalPages} 页，显示第
                                    {(currentPage - 1) * itemsPerPage + 1} - {Math.min(
                                        currentPage * itemsPerPage,
                                        matchResults.length,
                                    )} 条，总计 {matchResults.length} 条记录
                                </div>
                            {/if}
                        </div>
                    </div>
                {/if}
            {/if}
        </div>
    </div>
</div>
