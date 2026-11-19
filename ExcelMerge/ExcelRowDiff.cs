using System.Linq;
using System.Collections.Generic;

namespace ExcelMerge
{
    public class ExcelRowDiff
    {
        public int Index { get; private set; }
        public SortedDictionary<int, ExcelCellDiff> Cells { get; private set; }

        /// <summary>源文件侧的物理行号；该侧不存在此行时为 null。</summary>
        public int? SrcRowIndex { get; private set; }

        /// <summary>目标文件侧的物理行号；该侧不存在此行时为 null。</summary>
        public int? DstRowIndex { get; private set; }

        public bool ExistsInSrc { get { return SrcRowIndex.HasValue; } }
        public bool ExistsInDst { get { return DstRowIndex.HasValue; } }

        public ExcelRowDiff(int index)
        {
            Index = index;
            Cells = new SortedDictionary<int, ExcelCellDiff>();
        }

        /// <summary>记录该对齐行在两侧的真实物理行号（由 DiffCells 根据对齐结果设置）。</summary>
        internal void SetRows(ExcelRow srcRow, ExcelRow dstRow)
        {
            SrcRowIndex = srcRow != null ? (int?)srcRow.OriginalRowIndex : null;
            DstRowIndex = dstRow != null ? (int?)dstRow.OriginalRowIndex : null;
        }

        public ExcelCellDiff CreateCell(ExcelCell src, ExcelCell dst, int columnIndex, ExcelCellStatus status)
        {
            var cell = new ExcelCellDiff(columnIndex, Index, src, dst, status);
            Cells.Add(cell.ColumnIndex, cell);

            return cell;
        }

        public bool IsModified()
        {
            return Cells.Any(c => c.Value.Status != ExcelCellStatus.None);
        }

        /// <summary>该行只存在于 dst（目标侧）。空行也按真实存在性判断，不再依赖格子状态。</summary>
        public bool IsAdded()
        {
            if (SrcRowIndex.HasValue || DstRowIndex.HasValue)
                return DstRowIndex.HasValue && !SrcRowIndex.HasValue;

            // 兼容未设置物理行号的手工构造场景
            return Cells.Count > 0 && Cells.All(c => c.Value.Status == ExcelCellStatus.Added);
        }

        /// <summary>该行只存在于 src。空行也按真实存在性判断，不再依赖格子状态。</summary>
        public bool IsRemoved()
        {
            if (SrcRowIndex.HasValue || DstRowIndex.HasValue)
                return SrcRowIndex.HasValue && !DstRowIndex.HasValue;

            // 兼容未设置物理行号的手工构造场景
            return Cells.Count > 0 && Cells.All(c => c.Value.Status == ExcelCellStatus.Removed);
        }

        public int ModifiedCellCount
        {
            get { return Cells.Count(c => c.Value.Status != ExcelCellStatus.None); }
        }


        // TODO: Add row status field and implemnt UpdateStaus method.
    }
}
