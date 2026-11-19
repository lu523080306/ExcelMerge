using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelSheetDiff
    {
        public ExcelSheet SrcSheet { get; set; }
        public ExcelSheet DstSheet { get; set; }
        public SortedDictionary<int, ExcelRowDiff> Rows { get; private set; }

        /// <summary>对齐列 → 列状态（None/Inserted/Deleted），V2 结构级应用需要。</summary>
        public Dictionary<int, ExcelColumnStatus> ColumnStatusMap { get; set; }

        private bool? hasAnyChanges;
        /// <summary>该 sheet 对是否存在任何差异（供 sheet 状态下拉着色使用）。</summary>
        public bool HasAnyChanges
        {
            get
            {
                if (!hasAnyChanges.HasValue)
                    hasAnyChanges = CreateSummary().HasDiff;

                return hasAnyChanges.Value;
            }
        }

        public ExcelSheetDiff()
        {
            Rows = new SortedDictionary<int, ExcelRowDiff>();
        }

        public ExcelRowDiff CreateRow()
        {
            var row = new ExcelRowDiff(Rows.Any() ? Rows.Keys.Last() + 1 : 0);
            Rows.Add(row.Index, row);

            return row;
        }

        /// <summary>创建对齐行并记录两侧真实物理行（任一侧可为 null）。</summary>
        public ExcelRowDiff CreateRow(ExcelRow srcRow, ExcelRow dstRow)
        {
            var row = CreateRow();
            row.SetRows(srcRow, dstRow);

            return row;
        }

        public ExcelSheetDiffSummary CreateSummary()
        {
            var addedRowCount = 0;
            var removedRowCount = 0;
            var modifiedRowCount = 0;
            var modifiedCellCount = 0;
            foreach (var row in Rows)
            {
                if (row.Value.IsAdded())
                    addedRowCount++;
                else if (row.Value.IsRemoved())
                    removedRowCount++;

                if (row.Value.IsModified())
                    modifiedRowCount++;

                modifiedCellCount += row.Value.ModifiedCellCount;
            }

            return new ExcelSheetDiffSummary
            {
                AddedRowCount = addedRowCount,
                RemovedRowCount = removedRowCount,
                ModifiedRowCount = modifiedRowCount,
                ModifiedCellCount = modifiedCellCount,
            };
        }

        public override string ToString()
        {
            return $"{SrcSheet?.Name} {DstSheet?.Name}";
        }
    }
}
