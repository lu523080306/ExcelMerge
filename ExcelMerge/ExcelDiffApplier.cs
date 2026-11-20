using System;
using System.Collections.Generic;
using System.Linq;

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
    /// 把 ExcelSheetDiff 的差异应用到目标文件（值级 + 结构级行列增删）。
    ///
    /// 坐标原则（V3）：不使用"当前目标行"游标，所有写入一律使用 diff 中记录的真实物理坐标：
    ///   - 行：ExcelRowDiff.SrcRowIndex / DstRowIndex（读取时保留物理行号，空行/Null 行也算真实存在）
    ///   - 列：对齐列 → 目标物理列的静态映射（列结构增删后重新计算）
    ///
    /// 执行顺序（保证物理坐标在结构变更前有效）：
    ///   1. 列结构：删目标独有列（降序）、插对侧独有列（升序）
    ///   2. 行值：只写两侧都存在的行，按目标物理行号直接落值
    ///   3. 行结构：删目标独有行、插对侧独有行（按目标物理位置降序，插入锚点=下一个目标存在行）
    ///
    /// 已知局限：NPOI 列搬移不处理合并单元格/公式引用改写；插入行不继承样式。
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
                    var srcToDst = direction == ExcelDiffDirection.SrcToDst;

                    // 全量：所有结构列都同步
                    var plan = BuildColumnPlan(diff, srcToDst, null, true, null);
                    ExecuteColumnOps(target, plan, result);

                    // 值：只写两侧都存在的行（物理行号在行结构变更前有效）
                    WriteValues(target, diff, srcToDst, plan, null, result);

                    // 行结构：全部目标独有行删除 / 对侧独有行插入
                    ExecuteRowOps(target, diff, srcToDst, null, plan, result);

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
        /// 未选中的行完全不动；值按目标物理行号写入，行结构按物理位置降序执行。
        /// 只同步行：列只在选中行需要时补插（其余行留空），不删除/重排其他列。
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
                    var srcToDst = direction == ExcelDiffDirection.SrcToDst;
                    var selected = new HashSet<int>(diffRowIndexes ?? new int[0]);

                    // 选中行里出现的“目标缺失列”必须先补插，否则整行同步会丢掉这些格；
                    // 补插的列只在选中行写值，其余行保持空白。
                    var insertColumns = GetInsertColumnsForSelectedRows(diff, srcToDst, selected);
                    var plan = BuildColumnPlan(diff, srcToDst, null, insertColumns.Count > 0, insertColumns);

                    ExecuteColumnOps(target, plan, result);
                    WriteValues(target, diff, srcToDst, plan, selected, result);
                    ExecuteRowOps(target, diff, srcToDst, selected, plan, result);

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
        ///   结构列(对侧独有/目标独有) → 只对选中列做插入/删除；
        ///   值列 → 对每一行按目标物理行号把选中列的值从对侧写回。
        /// 只同步列：对侧独有行在目标里补插一行（只写选中列，其余列留空），
        /// 目标独有行的选中列写空——目标是“源列有几个单元格就落几个”。
        /// 列位置按"本批实际执行的列增删"重新计算。
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
                    var srcToDst = direction == ExcelDiffDirection.SrcToDst;
                    var selected = new HashSet<int>(alignedColumns ?? new int[0]);

                    var plan = BuildColumnPlan(diff, srcToDst, selected, true, selected);
                    ExecuteColumnOps(target, plan, result);

                    // 值：目标所有行的选中列都写（源侧缺行写空），WriteCol 已按选中集合过滤
                    WriteValues(target, diff, srcToDst, plan, null, result, includeTargetOnlyRows: true);

                    // 对侧独有行在目标里没有落点：补插一行，只写选中列，其他列留空
                    AlignRowsForColumnSync(target, diff, srcToDst, plan, result);

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
        /// 写值：默认只写两侧都存在的行；includeTargetOnlyRows=true 时（列同步）
        /// 目标独有行也写——源侧没有这一行，值即空串，用来把目标该列清空。
        /// 只存在于源侧的行没有目标落点，任何模式都不写（列同步不插行）。
        /// </summary>
        private static void WriteValues(
            ExcelWriteTarget target, ExcelSheetDiff diff, bool srcToDst,
            ColumnPlan plan, HashSet<int> selectedRows, ExcelApplyResult result,
            bool includeTargetOnlyRows = false)
        {
            foreach (var pair in diff.Rows)
            {
                if (selectedRows != null && !selectedRows.Contains(pair.Key))
                    continue;

                var row = pair.Value;

                var existsInTarget = srcToDst ? row.ExistsInDst : row.ExistsInSrc;
                var existsInOther = srcToDst ? row.ExistsInSrc : row.ExistsInDst;

                // 目标没有这一行 → 没有落点（列同步不插行）
                if (!existsInTarget)
                    continue;

                // 目标独有行：行同步/全量由行结构操作处理；列同步按“源侧无此行”写空
                if (!existsInOther && !includeTargetOnlyRows)
                    continue;

                var targetRow = (srcToDst ? row.DstRowIndex : row.SrcRowIndex).Value;

                foreach (var cellPair in row.Cells)
                {
                    var cell = cellPair.Value;

                    // 只写真正有差异的格：None(等值)格重写只会把目标文件
                    // 空格写成物理格子/撑宽行，重读时列对齐会把它判成幽灵差异。
                    if (cell.Status == ExcelCellStatus.None)
                        continue;

                    var col = plan.WriteCol[cellPair.Key];
                    if (col < 0)
                        continue;

                    target.SetCellValue(targetRow, col, srcToDst ? cell.SrcCell.Value : cell.DstCell.Value);
                    result.AppliedCellCount++;
                }
            }
        }

        /// <summary>
        /// 列同步时的行对齐：对侧独有行（源侧有、目标没有）补插一行，只写选中列的值，
        /// 其他列留空——保证“源列有几个单元格就落几个”。目标独有行保持不动
        /// （其选中列由 WriteValues 写空）。
        /// </summary>
        private static void AlignRowsForColumnSync(
            ExcelWriteTarget target, ExcelSheetDiff diff, bool srcToDst, ColumnPlan plan, ExcelApplyResult result)
        {
            var ops = new List<RowOp>();
            var wants = new List<ExcelRowDiff>();
            var gapAnchor = 0;

            foreach (var pair in diff.Rows)
            {
                var row = pair.Value;
                var inTarget = ExistsInTarget(row, srcToDst);
                var inOther = ExistsInOther(row, srcToDst);

                if (inTarget && inOther)
                {
                    FlushColumnGap(wants, gapAnchor, plan, srcToDst, ops);
                    gapAnchor = TargetRow(row, srcToDst) + 1;
                }
                else if (inTarget)
                {
                    // 目标独有行占位：插入锚点移到它后面
                    var physical = TargetRow(row, srcToDst);
                    if (physical + 1 > gapAnchor)
                        gapAnchor = physical + 1;
                }
                else if (inOther)
                {
                    wants.Add(row);
                }
            }

            FlushColumnGap(wants, gapAnchor, plan, srcToDst, ops);

            ops.Sort((a, b) =>
            {
                var cmp = b.Key.CompareTo(a.Key);
                if (cmp != 0)
                    return cmp;

                return b.Seq.CompareTo(a.Seq);
            });

            foreach (var op in ops)
            {
                target.InsertRow(op.Key, op.Values);
                result.InsertedRowCount++;
                result.AppliedCellCount += CountValues(op.Values);
            }
        }

        /// <summary>把空档里对侧独有行插到锚点位置（同锚点多行按对齐顺序逆序执行）。</summary>
        private static void FlushColumnGap(
            List<ExcelRowDiff> wants, int gapAnchor, ColumnPlan plan, bool srcToDst,
            List<RowOp> ops)
        {
            foreach (var want in wants)
            {
                // 只写选中列：其他列留空，列同步不应把整行都带过去
                var values = BuildInsertValues(want, srcToDst, plan, selectedColumnsOnly: true);
                if (!values.Any(v => !string.IsNullOrEmpty(v)))
                    continue;

                ops.Add(new RowOp
                {
                    Key = gapAnchor,
                    IsDelete = false,
                    Seq = want.Index,
                    Values = values,
                });
            }

            wants.Clear();
        }

        private static int CountValues(IReadOnlyList<string> values)
        {
            var count = 0;
            for (var i = 0; i < values.Count; i++)
            {
                if (!string.IsNullOrEmpty(values[i]))
                    count++;
            }

            return count;
        }

        private static void ExecuteColumnOps(ExcelWriteTarget target, ColumnPlan plan, ExcelApplyResult result)
        {
            var deletes = new List<ColumnOp>();
            var inserts = new List<ColumnOp>();
            foreach (var op in plan.Ops)
            {
                if (op.IsDelete) deletes.Add(op);
                else inserts.Add(op);
            }

            // 删除按操作前物理列降序：只影响更大的列号
            deletes.Sort((a, b) => b.PhysicalIndex.CompareTo(a.PhysicalIndex));
            // 插入按操作后物理列升序：位置已是最终布局
            inserts.Sort((a, b) => a.FinalIndex.CompareTo(b.FinalIndex));

            foreach (var op in deletes)
            {
                target.DeleteColumn(op.PhysicalIndex);
                result.DeletedColumnCount++;
            }

            foreach (var op in inserts)
            {
                target.InsertColumn(op.FinalIndex);
                result.InsertedColumnCount++;
            }
        }

        private static void ExecuteRowOps(
            ExcelWriteTarget target, ExcelSheetDiff diff, bool srcToDst,
            HashSet<int> selected, ColumnPlan plan, ExcelApplyResult result)
        {
            var ops = BuildRowOps(diff, srcToDst, selected, plan);
            foreach (var op in ops)
            {
                if (op.IsDelete)
                {
                    target.DeleteRow(op.Key);
                    result.DeletedRowCount++;
                }
                else
                {
                    target.InsertRow(op.Key, op.Values);
                    result.InsertedRowCount++;
                    result.AppliedCellCount += CountValues(op.Values);
                }
            }
        }

        /// <summary>
        /// 行结构操作单：按目标物理行号降序执行。
        /// 同一位置：先删后插；同一锚点的多行插入按对齐顺序逆序，保证源顺序。
        /// </summary>
        private static List<RowOp> BuildRowOps(
            ExcelSheetDiff diff, bool srcToDst, HashSet<int> selected, ColumnPlan plan)
        {
            var ops = new List<RowOp>();

            var maxTargetRow = -1;
            foreach (var pair in diff.Rows)
            {
                if (!ExistsInTarget(pair.Value, srcToDst))
                    continue;

                var physical = TargetRow(pair.Value, srcToDst);
                if (physical > maxTargetRow)
                    maxTargetRow = physical;
            }

            var appendAnchor = maxTargetRow + 1;

            // 逆序找插入锚点：下一个"目标侧存在"的行
            var keys = new List<int>(diff.Rows.Keys);
            int? nextTargetRow = null;
            for (var i = keys.Count - 1; i >= 0; i--)
            {
                var key = keys[i];
                var row = diff.Rows[key];

                if (ExistsInTarget(row, srcToDst))
                {
                    nextTargetRow = TargetRow(row, srcToDst);
                    continue;
                }

                if (!ExistsInOther(row, srcToDst))
                    continue;

                if (selected != null && !selected.Contains(key))
                    continue;

                ops.Add(new RowOp
                {
                    Key = nextTargetRow ?? appendAnchor,
                    IsDelete = false,
                    Seq = key,
                    Values = BuildInsertValues(row, srcToDst, plan),
                });
            }

            // 删除：目标侧独有
            foreach (var pair in diff.Rows)
            {
                var row = pair.Value;
                if (!ExistsInTarget(row, srcToDst) || ExistsInOther(row, srcToDst))
                    continue;

                if (selected != null && !selected.Contains(pair.Key))
                    continue;

                ops.Add(new RowOp
                {
                    Key = TargetRow(row, srcToDst),
                    IsDelete = true,
                    Seq = pair.Key,
                });
            }

            ops.Sort((a, b) =>
            {
                var cmp = b.Key.CompareTo(a.Key);
                if (cmp != 0)
                    return cmp;

                if (a.IsDelete != b.IsDelete)
                    return a.IsDelete ? -1 : 1;

                return b.Seq.CompareTo(a.Seq);
            });

            return ops;
        }

        /// <summary>插入行的值数组：下标为目标物理列（操作后布局）。</summary>
        /// <summary>
        /// 插入行的值数组：下标为目标物理列（操作后布局）。
        /// selectedColumnsOnly=true 时只填 plan.WriteCol 命中的列（列同步），
        /// 其余列留空，避免“同步一列却把整行都带过去”。
        /// </summary>
        private static string[] BuildInsertValues(
            ExcelRowDiff row, bool srcToDst, ColumnPlan plan, bool selectedColumnsOnly = false)
        {
            var values = new string[plan.Count < 0 ? 0 : plan.Count];
            for (var i = 0; i < values.Length; i++)
                values[i] = string.Empty;

            foreach (var pair in row.Cells)
            {
                var col = selectedColumnsOnly ? plan.WriteCol[pair.Key] : plan.PostOpCol[pair.Key];
                if (col < 0 || col >= values.Length)
                    continue;

                var cell = pair.Value;
                values[col] = srcToDst ? cell.SrcCell.Value : cell.DstCell.Value;
            }

            return values;
        }

        private static bool ExistsInTarget(ExcelRowDiff row, bool srcToDst)
        {
            return srcToDst ? row.ExistsInDst : row.ExistsInSrc;
        }

        private static bool ExistsInOther(ExcelRowDiff row, bool srcToDst)
        {
            return srcToDst ? row.ExistsInSrc : row.ExistsInDst;
        }

        private static int TargetRow(ExcelRowDiff row, bool srcToDst)
        {
            return (srcToDst ? row.DstRowIndex : row.SrcRowIndex).Value;
        }

        /// <summary>
        /// 选中行里，目标不存在的列（对侧独有）且该行有值的列。
        /// 行同步时先补插这些列，选中行的这些格才有落点（其余行留空）。
        /// </summary>
        private static HashSet<int> GetInsertColumnsForSelectedRows(
            ExcelSheetDiff diff, bool srcToDst, HashSet<int> selected)
        {
            var needed = new HashSet<int>();
            var status = diff.ColumnStatusMap;

            foreach (var pair in diff.Rows)
            {
                if (!selected.Contains(pair.Key))
                    continue;

                foreach (var cellPair in pair.Value.Cells)
                {
                    var c = cellPair.Key;
                    var s = status != null && status.ContainsKey(c) ? status[c] : ExcelColumnStatus.None;

                    // 对侧独有：srcToDst 时是 Deleted；反向时是 Inserted
                    var isOtherOnly = srcToDst
                        ? s == ExcelColumnStatus.Deleted
                        : s == ExcelColumnStatus.Inserted;
                    if (!isOtherOnly)
                        continue;

                    var cell = cellPair.Value;
                    var value = srcToDst ? cell.SrcCell.Value : cell.DstCell.Value;
                    if (!string.IsNullOrEmpty(value))
                        needed.Add(c);
                }
            }

            return needed;
        }

        private static int GetAlignedColumnCount(ExcelSheetDiff diff)
        {
            var count = 0;
            foreach (var row in diff.Rows.Values)
                count = Math.Max(count, row.Cells.Count);

            return count;
        }

        /// <summary>
        /// 计算目标侧"对齐列 → 物理列"的偏移。
        /// 目标文件可能被裁剪过前导空列，物理列 = 偏移 + 对齐列 - (目标侧缺失的对齐列数)。
        /// </summary>
        private static int GetTargetColumnOffset(
            ExcelSheetDiff diff, bool srcToDst, int[] delBefore, int[] insBefore)
        {
            foreach (var pair in diff.Rows)
            {
                foreach (var cellPair in pair.Value.Cells)
                {
                    var c = cellPair.Key;
                    var cell = cellPair.Value;

                    if (srcToDst)
                    {
                        if (cell.DstCell == null || cell.DstCell.IsSynthetic)
                            continue;

                        return cell.DstCell.OriginalColumnIndex - (c - delBefore[c]);
                    }

                    if (cell.SrcCell == null || cell.SrcCell.IsSynthetic)
                        continue;

                    return cell.SrcCell.OriginalColumnIndex - (c - insBefore[c]);
                }
            }

            return 0;
        }

        /// <summary>
        /// 列布局规划：
        ///   PostOpCol：结构操作后的目标物理列（-1=目标不存在）
        ///   WriteCol：值写入的目标物理列（-1=不写）
        ///   Ops：本批要执行的列结构操作（仅 selected 范围内的结构列）
        /// </summary>
        private static ColumnPlan BuildColumnPlan(
            ExcelSheetDiff diff, bool srcToDst, HashSet<int> selected,
            bool doStructuralOps, HashSet<int> structuralColumns)
        {
            var status = diff.ColumnStatusMap;
            var alignedCount = GetAlignedColumnCount(diff);
            if (status != null && status.Count > alignedCount)
                alignedCount = status.Count;

            var delBefore = new int[alignedCount + 1];
            var insBefore = new int[alignedCount + 1];
            for (var c = 0; c < alignedCount; c++)
            {
                var s = status != null && status.ContainsKey(c) ? status[c] : ExcelColumnStatus.None;
                delBefore[c + 1] = delBefore[c] + (s == ExcelColumnStatus.Deleted ? 1 : 0);
                insBefore[c + 1] = insBefore[c] + (s == ExcelColumnStatus.Inserted ? 1 : 0);
            }

            var offset = GetTargetColumnOffset(diff, srcToDst, delBefore, insBefore);

            var plan = new ColumnPlan();
            plan.PostOpCol = new int[alignedCount];
            plan.WriteCol = new int[alignedCount];
            for (var c = 0; c < alignedCount; c++)
            {
                plan.PostOpCol[c] = -1;
                plan.WriteCol[c] = -1;
            }

            var pos = offset;
            for (var c = 0; c < alignedCount; c++)
            {
                var s = status != null && status.ContainsKey(c) ? status[c] : ExcelColumnStatus.None;
                var selectedColumn = selected == null || selected.Contains(c);
                var structuralColumn = structuralColumns == null || structuralColumns.Contains(c);

                // 目标侧独有：srcToDst 时是 Inserted；反向时是 Deleted
                var isTargetOnly = srcToDst
                    ? s == ExcelColumnStatus.Inserted
                    : s == ExcelColumnStatus.Deleted;
                // 对侧独有：srcToDst 时是 Deleted；反向时是 Inserted
                var isOtherOnly = srcToDst
                    ? s == ExcelColumnStatus.Deleted
                    : s == ExcelColumnStatus.Inserted;

                if (s == ExcelColumnStatus.None)
                {
                    plan.PostOpCol[c] = pos;
                    plan.WriteCol[c] = selectedColumn ? pos : -1;
                    pos++;
                }
                else if (isTargetOnly)
                {
                    if (doStructuralOps && structuralColumn)
                    {
                        // 目标侧独有列 → 删除（操作前物理列）
                        var physical = offset + c - (srcToDst ? delBefore[c] : insBefore[c]);
                        plan.Ops.Add(new ColumnOp { IsDelete = true, PhysicalIndex = physical });
                    }
                    else
                    {
                        // 目标独有列：行同步时源侧没有这一列 → 写空；
                        // 列同步且未选中该列时保留目标原值（WriteCol=-1）
                        plan.PostOpCol[c] = pos;
                        plan.WriteCol[c] = selectedColumn ? pos : -1;
                        pos++;
                    }
                }
                else if (isOtherOnly)
                {
                    if (doStructuralOps && structuralColumn)
                    {
                        // 对侧独有列 → 在最终布局位置插入，并写入对侧值
                        plan.PostOpCol[c] = pos;
                        plan.WriteCol[c] = selectedColumn ? pos : -1;
                        plan.Ops.Add(new ColumnOp { IsDelete = false, FinalIndex = pos });
                        pos++;
                    }
                }
            }

            plan.Count = pos;
            return plan;
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

        private sealed class ColumnOp
        {
            public bool IsDelete;
            public int PhysicalIndex; // 删除：操作前物理列
            public int FinalIndex;    // 插入：操作后物理列
        }

        private sealed class ColumnPlan
        {
            public int[] PostOpCol;
            public int[] WriteCol;
            public int Count;
            public readonly List<ColumnOp> Ops = new List<ColumnOp>();
        }

        private sealed class RowOp
        {
            public int Key;         // 目标物理行
            public bool IsDelete;
            public int Seq;         // 对齐顺序
            public string[] Values; // 插入行值（下标=目标物理列）
        }
    }
}
