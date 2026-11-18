using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    /// <summary>
    /// 写回目标：对"目标文件"施加结构级(行列增删)与单元格值改动并保存。
    /// 行/列索引一律是"当前(活)模型"的真实坐标 —— 改动是边算边落，索引不靠缓存推导。
    /// CSV/TSV：内存网格等价实现；NPOI：直接操作 ISheet。
    /// 已知局限（结构级）：NPOI 列增删逐格搬移，不处理合并单元格/公式引用改写；插入行不继承样式。
    /// </summary>
    public abstract class ExcelWriteTarget : IDisposable
    {
        /// <summary>目标当前总行数。</summary>
        public abstract int RowCount { get; }

        /// <summary>读出指定格值（用于列结构搬移时保内容）。</summary>
        public abstract string GetCellValue(int row, int column);

        /// <summary>把指定格值写为 value（目标原为数字/布尔时保持类型）。</summary>
        public abstract void SetCellValue(int row, int column, string value);

        /// <summary>在第 rowIndex 处插入一行；values 按下标=最终列位置排列（空串=空单元格）。</summary>
        public abstract void InsertRow(int rowIndex, IReadOnlyList<string> values);

        /// <summary>删除第 rowIndex 行。</summary>
        public abstract void DeleteRow(int rowIndex);

        /// <summary>在第 columnIndex 处插入一空列（值由调用方随后逐行写入）。</summary>
        public abstract void InsertColumn(int columnIndex);

        /// <summary>删除第 columnIndex 列。</summary>
        public abstract void DeleteColumn(int columnIndex);

        /// <summary>把累积的改动落盘。</summary>
        public abstract void Save();

        public virtual void Dispose() { }

        public static ExcelWriteTarget Create(string path, string sheetName)
        {
            var ext = Path.GetExtension(path);
            if (ext == ".csv")
                return new SeparatedValuesTarget(path, ',');
            if (ext == ".tsv")
                return new SeparatedValuesTarget(path, '\t');

            return new NpoiWriteTarget(path, sheetName);
        }
    }

    /// <summary>NPOI 写回：打开原文件，结构/值改动直接落在 ISheet 上，最后整体写出（其余 sheet 保留）。</summary>
    internal sealed class NpoiWriteTarget : ExcelWriteTarget
    {
        private readonly string path;
        private readonly IWorkbook workbook;
        private readonly ISheet sheet;
        private bool[] numericColumnCache;
        private int numericColumnCacheSize;

        public NpoiWriteTarget(string path, string sheetName)
        {
            this.path = path;
            workbook = WorkbookFactory.Create(path);
            sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
                throw new ArgumentException("目标文件中找不到工作表: " + sheetName);
        }

        public override int RowCount
        {
            get { return sheet.LastRowNum + 1; }
        }

        public override string GetCellValue(int row, int column)
        {
            var npoiRow = sheet.GetRow(row);
            if (npoiRow == null)
                return string.Empty;

            var cell = npoiRow.GetCell(column);
            return ExcelUtility.GetCellStringValue(cell);
        }

        public override void SetCellValue(int row, int column, string value)
        {
            var npoiRow = sheet.GetRow(row) ?? sheet.CreateRow(row);
            var cell = npoiRow.GetCell(column) ?? npoiRow.CreateCell(column);
            ExcelUtility.SetValue(cell, value);
        }

        public override void InsertRow(int rowIndex, IReadOnlyList<string> values)
        {
            if (rowIndex <= sheet.LastRowNum)
                sheet.ShiftRows(rowIndex, sheet.LastRowNum, 1, false, false);

            var npoiRow = sheet.CreateRow(rowIndex);
            for (int i = 0; i < values.Count; i++)
            {
                var cell = npoiRow.CreateCell(i);
                var value = values[i];
                double d;
                if (ColumnHasNumeric(i) && double.TryParse(value, out d))
                    cell.SetCellValue(d);
                else
                    cell.SetCellValue(value);
            }
        }

        /// <summary>惰性探测某列在原始文件中是否为数字列（用于插入行保持数字类型）。</summary>
        private bool ColumnHasNumeric(int column)
        {
            EnsureNumericCache(column + 1);
            return column < numericColumnCache.Length && numericColumnCache[column];
        }

        private void EnsureNumericCache(int size)
        {
            if (numericColumnCache != null && numericColumnCacheSize >= size)
                return;

            numericColumnCache = new bool[Math.Max(size, (numericColumnCache == null ? 0 : numericColumnCache.Length))];
            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                var npoiRow = sheet.GetRow(r);
                if (npoiRow == null)
                    continue;

                for (int c = 0; c < npoiRow.LastCellNum && c < numericColumnCache.Length; c++)
                {
                    var cell = npoiRow.GetCell(c);
                    if (cell != null && cell.CellType == CellType.Numeric)
                        numericColumnCache[c] = true;
                }
            }

            numericColumnCacheSize = numericColumnCache.Length;
        }

        public override void DeleteRow(int rowIndex)
        {
            var npoiRow = sheet.GetRow(rowIndex);
            if (npoiRow != null)
                sheet.RemoveRow(npoiRow);

            if (rowIndex <= sheet.LastRowNum)
                sheet.ShiftRows(rowIndex, sheet.LastRowNum, -1);
        }

        public override void InsertColumn(int columnIndex)
        {
            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                var npoiRow = sheet.GetRow(r);
                if (npoiRow == null)
                    continue;

                var lastIndex = npoiRow.LastCellNum - 1;
                for (int c = lastIndex; c >= columnIndex && c >= 0; c--)
                {
                    var src = npoiRow.GetCell(c);
                    if (src == null)
                        continue;

                    var dst = npoiRow.GetCell(c + 1) ?? npoiRow.CreateCell(c + 1);
                    CopyCell(src, dst);
                    npoiRow.RemoveCell(src);   // 搬走后清掉原位，columnIndex 处留空给调用方写值
                }
            }
        }

        public override void DeleteColumn(int columnIndex)
        {
            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                var npoiRow = sheet.GetRow(r);
                if (npoiRow == null)
                    continue;

                var lastIndex = npoiRow.LastCellNum - 1;
                if (columnIndex > lastIndex)
                    continue;

                for (int c = columnIndex; c < lastIndex; c++)
                {
                    var src = npoiRow.GetCell(c + 1);
                    if (src != null)
                    {
                        var dst = npoiRow.GetCell(c) ?? npoiRow.CreateCell(c);
                        CopyCell(src, dst);
                    }
                    else
                    {
                        var hole = npoiRow.GetCell(c);
                        if (hole != null)
                            npoiRow.RemoveCell(hole);
                    }
                }

                var tail = npoiRow.GetCell(lastIndex);
                if (tail != null)
                    npoiRow.RemoveCell(tail);
            }
        }

        public override void Save()
        {
            // 先写临时文件，再释放读句柄后替换目标，避免"自己锁自己"（构造时打开的读句柄）
            var tmpPath = path + ".emtmp";
            using (var fs = new FileStream(tmpPath, FileMode.Create, FileAccess.Write))
                workbook.Write(fs);

            ExcelUtility.ReleaseWorkbook(workbook);
            File.Delete(path);
            File.Move(tmpPath, path);
        }

        /// <summary>复制单元格（值+类型+公式），丢弃样式。</summary>
        private static void CopyCell(ICell src, ICell dst)
        {
            switch (src.CellType)
            {
                case CellType.Numeric:
                    dst.SetCellValue(src.NumericCellValue);
                    break;
                case CellType.Boolean:
                    dst.SetCellValue(src.BooleanCellValue);
                    break;
                case CellType.Formula:
                    dst.SetCellFormula(src.CellFormula);
                    break;
                case CellType.String:
                    dst.SetCellValue(src.StringCellValue);
                    break;
                default:
                    dst.SetCellValue(ExcelUtility.GetCellStringValue(src));
                    break;
            }
        }
    }

    /// <summary>CSV/TSV 写回：与本库读取器对称的朴素分割/拼接（UTF-8 无 BOM）。</summary>
    internal sealed class SeparatedValuesTarget : ExcelWriteTarget
    {
        private readonly string path;
        private readonly char delimiter;
        private readonly List<string[]> grid;
        private bool saved;

        public SeparatedValuesTarget(string path, char delimiter)
        {
            this.path = path;
            this.delimiter = delimiter;
            grid = new List<string[]>();

            foreach (var line in File.ReadAllLines(path, Encoding.UTF8))
                grid.Add(line.Split(delimiter));
        }

        public override int RowCount
        {
            get { return grid.Count; }
        }

        public override string GetCellValue(int row, int column)
        {
            if (row < 0 || row >= grid.Count)
                return string.Empty;

            var cells = grid[row];
            return column >= 0 && column < cells.Length ? cells[column] : string.Empty;
        }

        public override void SetCellValue(int row, int column, string value)
        {
            EnsureRow(row);
            var cells = grid[row];
            if (column > cells.Length)
                Array.Resize(ref cells, column + 1);

            cells[column] = value;
            grid[row] = cells;
        }

        public override void InsertRow(int rowIndex, IReadOnlyList<string> values)
        {
            grid.Insert(Math.Min(rowIndex, grid.Count), values == null ? new string[0] : new List<string>(values).ToArray());
        }

        public override void DeleteRow(int rowIndex)
        {
            if (rowIndex >= 0 && rowIndex < grid.Count)
                grid.RemoveAt(rowIndex);
        }

        public override void InsertColumn(int columnIndex)
        {
            for (int r = 0; r < grid.Count; r++)
            {
                var cells = grid[r];
                var newCells = new string[Math.Max(cells.Length, columnIndex) + 1];
                for (int c = 0, i = 0; i < newCells.Length; i++)
                {
                    if (i == columnIndex)
                        continue;

                    newCells[i] = c < cells.Length ? cells[c] : string.Empty;
                    c++;
                }

                grid[r] = newCells;
            }
        }

        public override void DeleteColumn(int columnIndex)
        {
            for (int r = 0; r < grid.Count; r++)
            {
                var cells = grid[r];
                if (columnIndex >= cells.Length)
                    continue;

                var newCells = new string[cells.Length - 1];
                for (int c = 0, i = 0; c < cells.Length; c++)
                {
                    if (c == columnIndex)
                        continue;

                    newCells[i++] = cells[c];
                }

                grid[r] = newCells;
            }
        }

        public override void Save()
        {
            if (saved)
                return;

            using (var writer = new StreamWriter(path, false, new UTF8Encoding(false)))
            {
                foreach (var line in grid)
                    writer.WriteLine(string.Join(delimiter.ToString(), line));

                writer.Flush();
            }

            saved = true;
        }

        private void EnsureRow(int row)
        {
            while (grid.Count <= row)
                grid.Add(new string[0]);
        }
    }
}