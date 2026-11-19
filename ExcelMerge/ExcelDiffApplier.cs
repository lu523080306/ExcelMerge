using System;
using System.Collections.Generic;

namespace ExcelMerge
{
    public enum ExcelDiffDirection
    {
        SrcToDst, // 把 src 的结构与值应用到 dst 文件（让 dst 与 src 对齐）
        DstToSrc, // 把 dst 的结构与值应用到 src 文件
    }

    /// <summary>一次单元格值变更（以"目标文件"的真实坐标表达）。</summary>
    public sealed class CellChange
    {
        public int Row { get; }
        public int Column { get; }
        public string Value { get; }

        public CellChange(int row, int column, string value)
        {
            Row = row;
            Column = column;
            Value = value;
        }
    }

    public sealed class ExcelApplyResult
    {
        public bool IsSuccess { get; set; }
        public int AppliedCellCount { get; set; }
        public int InsertedRowCount { get; set; }
        public int DeletedRowCount { get; set; }
        public int InsertedColumnCount { get; set; }
        public int DeletedColumnCount { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// 把 ExcelSheetDiff 的差异应用到目标文件（V2：值级 + 结构级行列增删）。
    ///
    /// 算法分三步（全部作用于活模型，索引实时计算）：
    ///   1. 列结构同步：删目标独有列、插对侧独有列（先删后插，索引在删/插当下仍有效）。
    ///   2. 行级前向扫描：对目标侧本有的行写入"对侧值"（等值格写相同值=无操作）；
    ///      对目标侧独有的行删除；对仅在对侧的行插入。
    ///   3. 保存。
    ///
    /// 已知前提：行结构应用假设目标行连续（配置表典型形态）；NPOI 列搬移不处理合并单元格与公式引用改写。
    /// </summary>
    public class ExcelDiffApplier
    {
        public ExcelApplyResult Apply(
            ExcelSheetDiff diff, string targetFilePath, string targetSheetName, ExcelDiffDirection direction)
        {
            var result = new ExcelApplyResult();

            try
            {
                using (var target = ExcelWriteTarget.Create(targetFilePath, targetSheetName))
                {
                    var status = diff.ColumnStatusMap;
                    var alignedCount = status == null || status.Count == 0
                        ? GetAlignedColumnCount(diff)
                        : status.Count;

                    // 目标侧要移除的列状态：让 dst 对齐 src 时移除 Inserted(仅dst)；反向移除 Deleted(仅src)
                    var removeStatus = direction == ExcelDiffDirection.SrcToDst
                        ? ExcelColumnStatus.Inserted
                        : ExcelColumnStatus.Deleted;

                    int[] delBefore; int[] insBefore; int[] finalCol;
                    BuildColumnMaps(status, alignedCount, removeStatus, out delBefore, out insBefore, out finalCol);

                    SyncColumns(target, status, alignedCount, direction, delBefore, insBefore, finalCol, result);

                    SyncRowsAndValues(target, diff, direction, finalCol, result);

                    target.Save();

                    result.IsSuccess = true;
                    result.Message = BuildMessage(direction, result);
                }

                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.Message = "应用失败: " + ex.Message;
                return result;
            }
        }

        /// <summary>显式指定一组单元格改动（"选中格应用"模式，不做行列结构同步）。</summary>
        public ExcelApplyResult ApplyExplicit(string targetFilePath, string targetSheetName, ICollection<CellChange> changes)
        {
            var result = new ExcelApplyResult();

            try
            {
                using (var target = ExcelWriteTarget.Create(targetFilePath, targetSheetName))
                {
                    foreach (var c in changes)
                    {
                        target.SetCellValue(c.Row, c.Column, c.Value);
                        result.AppliedCellCount++;
                    }

                    target.Save();

                    result.IsSuccess = true;
                    result.Message = "已应用 " + changes.Count + " 个选中格子的值";
                }

                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.Message = "应用失败: " + ex.Message;
                return result;
            }
        }

        /// <summary>
        /// 只应用选中的 diff 行（"同步一行"模式）：结构行(Added/Removed)做插入/删除，普通行做整行值回写。
        /// 未选中的行跳过（保持目标现状），游标仍完整走查，保证被选行的坐标对齐。
        /// 不做列结构同步；列位置按最终对齐布局映射。
        /// </summary>
        public ExcelApplyResult ApplyRows(
            ExcelSheetDiff diff, string targetFilePath, string targetSheetName,
            ExcelDiffDirection direction, ICollection<int> diffRowIndexes)
        {
            var result = new ExcelApplyResult();

            try
            {
                using (var target = ExcelWriteTarget.Create(targetFilePath, targetSheetName))
                {
                    var selected = new HashSet<int>(diffRowIndexes ?? new int[0]);
                    var alignedCount = GetAlignedColumnCount(diff);
                    var removeStatus = direction == ExcelDiffDirection.SrcToDst
                        ? ExcelColumnStatus.Inserted
                        : ExcelColumnStatus.Deleted;

                    int[] delBefore; int[] insBefore; int[] finalCol;
                    BuildColumnMaps(diff.ColumnStatusMap, alignedCount, removeStatus, out delBefore, out insBefore, out finalCol);

                    var srcToDst = direction == ExcelDiffDirection.SrcToDst;
                    var cur = 0;

                    foreach (var pair in diff.Rows)
                    {
                        var row = pair.Value;
                        var isSelected = selected.Contains(pair.Key);
                        var hasCells = row.Cells.Count > 0;
                        var isSrcOnly = hasCells && row.IsRemoved();
                        var isDstOnly = hasCells && row.IsAdded();

                        if (srcToDst)
                        {
                            if (isSrcOnly)
                            {
                                if (isSelected) { target.InsertRow(cur, BuildRowValues(row, true, finalCol)); result.InsertedRowCount++; cur++; }
                                // 未选中：目标里本无此行，游标不动
                            }
                            else if (isDstOnly)
                            {
                                if (isSelected) { target.DeleteRow(cur); result.DeletedRowCount++; }
                                else cur++;   // 未选中：行保留在目标里，占一个位置
                            }
                            else
                            {
                                if (isSelected) WriteRowValues(target, row, cur, true, finalCol, result);
                                cur++;
                            }
                        }
                        else
                        {
                            if (isDstOnly)
                            {
                                if (isSelected) { target.InsertRow(cur, BuildRowValues(row, false, finalCol)); result.InsertedRowCount++; cur++; }
                            }
                            else if (isSrcOnly)
                            {
                                if (isSelected) { target.DeleteRow(cur); result.DeletedRowCount++; }
                                else cur++;   // 未选中：行保留在目标源文件里
                            }
                            else
                            {
                                if (isSelected) WriteRowValues(target, row, cur, false, finalCol, result);
                                cur++;
                            }
                        }
                    }

                    target.Save();

                    result.IsSuccess = true;
                    result.Message = BuildMessage(direction, result);
                }

                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.Message = "应用失败: " + ex.Message;
                return result;
            }
        }

        /// <summary>
        /// 只同步选中的对齐列（"同步一列"模式）：
        ///   结构列(对侧独有/目标独有) → 对该列做插入/删除；
        ///   值列 → 沿行游标把该列在每个目标行的值从对侧写回。
        /// 不做行结构同步；先结构列(删降/插升)再值列，保证布局自洽。
        /// </summary>
        public ExcelApplyResult ApplyColumns(
            ExcelSheetDiff diff, string targetFilePath, string targetSheetName,
            ExcelDiffDirection direction, ICollection<int> alignedColumns)
        {
            var result = new ExcelApplyResult();

            try
            {
                using (var target = ExcelWriteTarget.Create(targetFilePath, targetSheetName))
                {
                    var selected = new HashSet<int>(alignedColumns ?? new int[0]);
                    var alignedCount = GetAlignedColumnCount(diff);
                    var srcToDst = direction == ExcelDiffDirection.SrcToDst;
                    var removeStatus = srcToDst ? ExcelColumnStatus.Inserted : ExcelColumnStatus.Deleted;

                    int[] delBefore; int[] insBefore; int[] finalCol;
                    BuildColumnMaps(diff.ColumnStatusMap, alignedCount, removeStatus, out delBefore, out insBefore, out finalCol);
                    var colStatus = diff.ColumnStatusMap;

                    // 1) 结构列：删目标独有列(降序)，插对侧独有列(升序)——只对选中列
                    for (int c = alignedCount - 1; c >= 0; c--)
                    {
                        if (!selected.Contains(c))
                            continue;

                        var s = colStatus != null && colStatus.ContainsKey(c) ? colStatus[c] : ExcelColumnStatus.None;
                        var isTargetOnly = srcToDst ? s == ExcelColumnStatus.Inserted : s == ExcelColumnStatus.Deleted;
                        if (!isTargetOnly)
                            continue;

                        var originIndex = srcToDst ? c - delBefore[c] : c - insBefore[c];
                        target.DeleteColumn(originIndex);
                        result.DeletedColumnCount++;
                    }

                    var otherOnly = srcToDst ? ExcelColumnStatus.Deleted : ExcelColumnStatus.Inserted;
                    for (int c = 0; c < alignedCount; c++)
                    {
                        if (!selected.Contains(c))
                            continue;

                        var s = colStatus != null && colStatus.ContainsKey(c) ? colStatus[c] : ExcelColumnStatus.None;
                        if (s != otherOnly)
                            continue;

                        target.InsertColumn(finalCol[c]);
                        result.InsertedColumnCount++;
                    }

                    // 2) 值列：沿行游标，对目标本有的每一行，把选中列的值从对侧写回
                    var cur = 0;
                    foreach (var pair in diff.Rows)
                    {
                        var row = pair.Value;
                        var hasCells = row.Cells.Count > 0;
                        var isSrcOnly = hasCells && row.IsRemoved();
                        var isDstOnly = hasCells && row.IsAdded();

                        if (srcToDst)
                        {
                            if (isSrcOnly)
                            {
                                // 目标无此行，游标不动
                            }
                            else if (isDstOnly)
                            {
                                cur++;   // 目标本有，未做行同步 → 保留占位
                            }
                            else
                            {
                                WriteSelectedColumnValues(target, row, cur, true, finalCol, selected, result);
                                cur++;
                            }
                        }
                        else
                        {
                            if (isDstOnly)
                            {
                                // 目标(src)无此行，游标不动
                            }
                            else if (isSrcOnly)
                            {
                                cur++;   // 目标本有，保留占位
                            }
                            else
                            {
                                WriteSelectedColumnValues(target, row, cur, false, finalCol, selected, result);
                                cur++;
                            }
                        }
                    }

                    target.Save();

                    result.IsSuccess = true;
                    result.Message = BuildMessage(direction, result);
                }

                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.Message = "应用失败: " + ex.Message;
                return result;
            }
        }

        private static void WriteSelectedColumnValues(
            ExcelWriteTarget target, ExcelRowDiff row, int targetRow,
            bool fromSrc, int[] finalCol, HashSet<int> selected, ExcelApplyResult result)
        {
            foreach (var c in selected)
            {
                if (finalCol[c] < 0)
                    continue;

                ExcelCellDiff cell;
                if (!row.Cells.TryGetValue(c, out cell))
                    continue;

                // 只写有差异的格：None(等值)格改写会制造幽灵物理格
                if (cell.Status == ExcelCellStatus.None)
                    continue;

                target.SetCellValue(targetRow, finalCol[c], fromSrc ? cell.SrcCell.Value : cell.DstCell.Value);
                result.AppliedCellCount++;
            }
        }

        private static int GetAlignedColumnCount(ExcelSheetDiff diff)
        {
            var count = 0;
            foreach (var row in diff.Rows.Values)
                count = Math.Max(count, row.Cells.Count);

            return count;
        }

        private static void BuildColumnMaps(
            Dictionary<int, ExcelColumnStatus> status, int alignedCount, ExcelColumnStatus removeStatus,
            out int[] delBefore, out int[] insBefore, out int[] finalCol)
        {
            delBefore = new int[alignedCount + 1];
            insBefore = new int[alignedCount + 1];

            for (int c = 0; c < alignedCount; c++)
            {
                var s = status != null && status.ContainsKey(c) ? status[c] : ExcelColumnStatus.None;
                delBefore[c + 1] = delBefore[c] + (s == ExcelColumnStatus.Deleted ? 1 : 0);
                insBefore[c + 1] = insBefore[c] + (s == ExcelColumnStatus.Inserted ? 1 : 0);
            }

            finalCol = new int[alignedCount];
            for (int i = 0; i < alignedCount; i++)
                finalCol[i] = -1;

            var idx = 0;
            for (int c = 0; c < alignedCount; c++)
            {
                var s = status != null && status.ContainsKey(c) ? status[c] : ExcelColumnStatus.None;
                if (s != removeStatus)
                    finalCol[c] = idx++;
            }
        }

        /// <summary>列结构同步：先删目标独有的列（降序），再插对侧独有的列（升序，位置=最终布局）。</summary>
        private static void SyncColumns(
            ExcelWriteTarget target, Dictionary<int, ExcelColumnStatus> status, int alignedCount,
            ExcelDiffDirection direction, int[] delBefore, int[] insBefore, int[] finalCol, ExcelApplyResult result)
        {
            var srcToDst = direction == ExcelDiffDirection.SrcToDst;

            // 1) 删除目标独有的列
            for (int c = alignedCount - 1; c >= 0; c--)
            {
                var s = status != null && status.ContainsKey(c) ? status[c] : ExcelColumnStatus.None;
                var isTargetOnly = srcToDst ? s == ExcelColumnStatus.Inserted : s == ExcelColumnStatus.Deleted;
                if (!isTargetOnly)
                    continue;

                var originIndex = srcToDst ? c - delBefore[c] : c - insBefore[c];
                target.DeleteColumn(originIndex);
                result.DeletedColumnCount++;
            }

            // 2) 插入对侧独有的列
            var otherOnly = srcToDst ? ExcelColumnStatus.Deleted : ExcelColumnStatus.Inserted;
            for (int c = 0; c < alignedCount; c++)
            {
                var s = status != null && status.ContainsKey(c) ? status[c] : ExcelColumnStatus.None;
                if (s != otherOnly)
                    continue;

                target.InsertColumn(finalCol[c]);
                result.InsertedColumnCount++;
            }
        }

        /// <summary>行级前向扫描 + 值写入。</summary>
        private static void SyncRowsAndValues(
            ExcelWriteTarget target, ExcelSheetDiff diff, ExcelDiffDirection direction, int[] finalCol, ExcelApplyResult result)
        {
            var srcToDst = direction == ExcelDiffDirection.SrcToDst;
            var cur = 0;

            foreach (var rowPair in diff.Rows)
            {
                var row = rowPair.Value;
                var hasCells = row.Cells.Count > 0;
                var isSrcOnly = hasCells && row.IsRemoved();
                var isDstOnly = hasCells && row.IsAdded();

                if (srcToDst)
                {
                    if (isSrcOnly)
                    {
                        target.InsertRow(cur, BuildRowValues(row, true, finalCol));
                        result.InsertedRowCount++;
                        cur++;
                    }
                    else if (isDstOnly)
                    {
                        target.DeleteRow(cur);
                        result.DeletedRowCount++;
                    }
                    else
                    {
                        WriteRowValues(target, row, cur, true, finalCol, result);
                        cur++;
                    }
                }
                else
                {
                    if (isDstOnly)
                    {
                        target.InsertRow(cur, BuildRowValues(row, false, finalCol));
                        result.InsertedRowCount++;
                        cur++;
                    }
                    else if (isSrcOnly)
                    {
                        target.DeleteRow(cur);
                        result.DeletedRowCount++;
                    }
                    else
                    {
                        WriteRowValues(target, row, cur, false, finalCol, result);
                        cur++;
                    }
                }
            }
        }

        private static string[] BuildRowValues(ExcelRowDiff row, bool fromSrc, int[] finalCol)
        {
            var values = new string[finalCol.Length];
            for (int c = 0; c < finalCol.Length; c++)
            {
                if (finalCol[c] < 0)
                    continue;

                ExcelCellDiff cell;
                values[finalCol[c]] = row.Cells.TryGetValue(c, out cell)
                    ? (fromSrc ? cell.SrcCell.Value : cell.DstCell.Value)
                    : string.Empty;
            }

            return values;
        }

        private static void WriteRowValues(
            ExcelWriteTarget target, ExcelRowDiff row, int targetRow, bool fromSrc, int[] finalCol, ExcelApplyResult result)
        {
            for (int c = 0; c < finalCol.Length; c++)
            {
                if (finalCol[c] < 0)
                    continue;

                ExcelCellDiff cell;
                if (!row.Cells.TryGetValue(c, out cell))
                    continue;

                // 只写"真正需要同步"的格：None(等值)格 src==dst，重写只会把目标文件
                // 空格写成物理格子/撑宽行，重读时列对齐会把它判成幽灵差异。
                if (cell.Status == ExcelCellStatus.None)
                    continue;

                target.SetCellValue(targetRow, finalCol[c], fromSrc ? cell.SrcCell.Value : cell.DstCell.Value);
                result.AppliedCellCount++;
            }
        }

        private static string BuildMessage(ExcelDiffDirection direction, ExcelApplyResult result)
        {
            var parts = new List<string>();
            if (result.AppliedCellCount > 0)
                parts.Add(result.AppliedCellCount + " 个单元格");
            if (result.InsertedRowCount > 0)
                parts.Add("插入 " + result.InsertedRowCount + " 行");
            if (result.DeletedRowCount > 0)
                parts.Add("删除 " + result.DeletedRowCount + " 行");
            if (result.InsertedColumnCount > 0)
                parts.Add("插入 " + result.InsertedColumnCount + " 列");
            if (result.DeletedColumnCount > 0)
                parts.Add("删除 " + result.DeletedColumnCount + " 列");

            var text = parts.Count == 0
                ? "没有需要应用的差异"
                : "已应用: " + string.Join(", ", parts);

            return text + " (" + (direction == ExcelDiffDirection.SrcToDst ? "src→dst" : "dst→src") + ")";
        }
    }
}