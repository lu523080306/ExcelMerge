using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(ISheet sheet)
        {
            var actualRowIndex = 0;
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    // 中间的空行（null 行）也必须保留：否则模型行号被压缩，
                    // 后面的数据会顶到空行位置，与 Excel 行号及另一侧 diff 行整体错位。
                    // 物理行号仍记在 OriginalRowIndex，写回时用它定位。
                    yield return new ExcelRow(actualRowIndex++, rowIndex, new List<ExcelCell>());
                    continue;
                }

                var cells = new List<ExcelCell>();
                for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    var stringValue = ExcelUtility.GetCellStringValue(cell);

                    cells.Add(new ExcelCell(stringValue, columnIndex, rowIndex));
                }

                yield return new ExcelRow(actualRowIndex++, rowIndex, cells);
            }
        }
    }
}
