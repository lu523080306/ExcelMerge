using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using ExcelMerge;
using NPOI.XSSF.UserModel;

// 无头回归：ExcelMerge 核心库的 apply 功能（V2：值级 + 结构级行列增删）
// 场景：CSV 行增删双向、CSV 列增删双向、XLSX 行增删 + 数值类型保持
class ApplyVerify
{
    static int fails = 0;

    static void Check(string name, bool ok, string detail)
    {
        Console.WriteLine((ok ? "  [OK]  " : "  [FAIL] ") + name + (ok ? "" : "  ==> " + detail));
        if (!ok) fails++;
    }

    static ExcelSheet ReadSheet(string path, string sheetName)
    {
        return ExcelWorkbook.Create(path, new ExcelSheetReadConfig()).Sheets[sheetName];
    }

    static ExcelSheetDiff DoDiff(string pathA, string sheetA, string pathB, string sheetB)
    {
        return ExcelSheet.Diff(ReadSheet(pathA, sheetA), ReadSheet(pathB, sheetB), new ExcelSheetDiffConfig());
    }

    static string Metrics(string pathA, string sheetA, string pathB, string sheetB)
    {
        var diff = DoDiff(pathA, sheetA, pathB, sheetB);
        int mod = 0, add = 0, rem = 0, addRow = 0, remRow = 0;
        foreach (var rd in diff.Rows)
        {
            if (rd.Value.IsAdded()) addRow++;
            else if (rd.Value.IsRemoved()) remRow++;
            foreach (var c in rd.Value.Cells)
            {
                if (c.Value.Status == ExcelCellStatus.Modified) mod++;
                else if (c.Value.Status == ExcelCellStatus.Added) add++;
                else if (c.Value.Status == ExcelCellStatus.Removed) rem++;
            }
        }
        return "mod=" + mod + " add=" + add + " rem=" + rem + " addRow=" + addRow + " remRow=" + remRow;
    }

    static void CheckEqualFile(string pathA, string pathB)
    {
        var a = File.ReadAllLines(pathA, Encoding.UTF8);
        var b = File.ReadAllLines(pathB, Encoding.UTF8);
        bool ok = a.Length == b.Length;
        string detail = "行数 " + a.Length + " vs " + b.Length;
        if (ok)
            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i]) { ok = false; detail = "第" + (i + 1) + "行: [" + a[i] + "] vs [" + b[i] + "]"; break; }
        Check("文件等价 " + Path.GetFileName(pathA) + " === " + Path.GetFileName(pathB), ok, detail);
    }

    // ============ 场景1+2：CSV 行增删 + 值，双向 ============
    static void CsvRowSync(string dir)
    {
        Console.WriteLine("== [场景] CSV 行级结构 + 值 (a.csv vs b.csv，b 含新增行5 + 修改格) ==");
        var srcA = Path.Combine(dir, "a.csv");
        var srcB = Path.Combine(dir, "b.csv");

        // src→dst：让 b 对齐 a
        var t = Path.Combine(dir, "v2_b_to_a.csv");
        File.Copy(srcB, t, true);
        var r1 = new ExcelDiffApplier().Apply(DoDiff(srcA, "a.csv", t, "v2_b_to_a.csv"), t, "v2_b_to_a.csv", ExcelDiffDirection.SrcToDst);
        Check("src→dst 成功", r1.IsSuccess, r1.Message);
        Check("src→dst 删除 1 行(新增行5)", r1.DeletedRowCount == 1, "deleted=" + r1.DeletedRowCount);
        Check("src→dst 值应用 ≥4", r1.AppliedCellCount >= 4, "cells=" + r1.AppliedCellCount);
        Console.WriteLine("      指标(前) " + Metrics(srcA, "a.csv", t, "v2_b_to_a.csv"));
        Check("src→dst 后重对比=0差异", Metrics(srcA, "a.csv", t, "v2_b_to_a.csv") == "mod=0 add=0 rem=0 addRow=0 remRow=0", Metrics(srcA, "a.csv", t, "v2_b_to_a.csv"));
        CheckEqualFile(srcA, t);

        // dst→src：让 a 对齐 b
        t = Path.Combine(dir, "v2_a_to_b.csv");
        File.Copy(srcA, t, true);
        var r2 = new ExcelDiffApplier().Apply(DoDiff(t, "v2_a_to_b.csv", srcB, "b.csv"), t, "v2_a_to_b.csv", ExcelDiffDirection.DstToSrc);
        Check("dst→src 成功", r2.IsSuccess, r2.Message);
        Check("dst→src 插入 1 行(新增行5)", r2.InsertedRowCount == 1, "inserted=" + r2.InsertedRowCount);
        Console.WriteLine("      指标(前) " + Metrics(srcB, "b.csv", t, "v2_a_to_b.csv"));
        Check("dst→src 后重对比=0差异", Metrics(srcB, "b.csv", t, "v2_a_to_b.csv") == "mod=0 add=0 rem=0 addRow=0 remRow=0", Metrics(srcB, "b.csv", t, "v2_a_to_b.csv"));
        CheckEqualFile(srcB, t);
    }

    // ============ 场景3+4：CSV 行删除双向 ============
    static void CsvRowDelete(string dir)
    {
        Console.WriteLine("== [场景] CSV 行删除双向 (a vs c，c 少了行2) ==");
        var srcA = Path.Combine(dir, "a.csv");
        var srcC = Path.Combine(dir, "c.csv");

        // dst→src：让 c 补上行2
        var t = Path.Combine(dir, "v2_c_gain.csv");
        File.Copy(srcC, t, true);
        var r1 = new ExcelDiffApplier().Apply(DoDiff(t, "v2_c_gain.csv", srcA, "a.csv"), t, "v2_c_gain.csv", ExcelDiffDirection.DstToSrc);
        Check("c 补行2 成功", r1.IsSuccess && r1.InsertedRowCount == 1, r1.Message);
        Check("c 补行2 后与 a 等价", Metrics(srcA, "a.csv", t, "v2_c_gain.csv") == "mod=0 add=0 rem=0 addRow=0 remRow=0", Metrics(srcA, "a.csv", t, "v2_c_gain.csv"));
        CheckEqualFile(srcA, t);

        // src→dst：让 a 删掉行2
        t = Path.Combine(dir, "v2_a_lose.csv");
        File.Copy(srcA, t, true);
        var r2 = new ExcelDiffApplier().Apply(DoDiff(srcC, "c.csv", t, "v2_a_lose.csv"), t, "v2_a_lose.csv", ExcelDiffDirection.SrcToDst);
        Check("a 删行2 成功", r2.IsSuccess && r2.DeletedRowCount == 1, r2.Message);
        Check("a 删行2 后与 c 等价", Metrics(srcC, "c.csv", t, "v2_a_lose.csv") == "mod=0 add=0 rem=0 addRow=0 remRow=0", Metrics(srcC, "c.csv", t, "v2_a_lose.csv"));
        CheckEqualFile(srcC, t);
    }

    // ============ 场景5+6：CSV 列增删双向 ============
    static void CsvColumnSync(string dir)
    {
        Console.WriteLine("== [场景] CSV 列增删双向 (a vs wide，wide 多一列 note) ==");
        var srcA = Path.Combine(dir, "a.csv");
        var srcW = Path.Combine(dir, "wide.csv");

        // src→dst：让 wide 删掉 note 列
        var t = Path.Combine(dir, "v2_wide_lose_note.csv");
        File.Copy(srcW, t, true);
        var r1 = new ExcelDiffApplier().Apply(DoDiff(srcA, "a.csv", t, "v2_wide_lose_note.csv"), t, "v2_wide_lose_note.csv", ExcelDiffDirection.SrcToDst);
        Check("wide 删 note 列成功", r1.IsSuccess && r1.DeletedColumnCount == 1, r1.Message);
        Check("wide 删 note 后与 a 等价", Metrics(srcA, "a.csv", t, "v2_wide_lose_note.csv") == "mod=0 add=0 rem=0 addRow=0 remRow=0", Metrics(srcA, "a.csv", t, "v2_wide_lose_note.csv"));
        CheckEqualFile(srcA, t);

        // dst→src：让 a 补上 note 列(带值)
        t = Path.Combine(dir, "v2_a_gain_note.csv");
        File.Copy(srcA, t, true);
        var r2 = new ExcelDiffApplier().Apply(DoDiff(t, "v2_a_gain_note.csv", srcW, "wide.csv"), t, "v2_a_gain_note.csv", ExcelDiffDirection.DstToSrc);
        Check("a 补 note 列成功", r2.IsSuccess && r2.InsertedColumnCount == 1, r2.Message);
        Check("a 补 note 后与 wide 等价", Metrics(srcW, "wide.csv", t, "v2_a_gain_note.csv") == "mod=0 add=0 rem=0 addRow=0 remRow=0", Metrics(srcW, "wide.csv", t, "v2_a_gain_note.csv"));
        CheckEqualFile(srcW, t);
    }

    // ============ 场景7：XLSX 行增删 + 数值类型保持 ============
    static void XlsxRowAndType(string dir)
    {
        Console.WriteLine("== [场景] XLSX 行增删 + 数值类型保持 ==");
        var baseX = Path.Combine(dir, "x_base.xlsx");
        var wideX = Path.Combine(dir, "x_wide.xlsx");
        WriteXlsx(baseX, new string[] { "1,织田信长,100,50", "2,丰臣秀吉,90,40", "3,德川家康,80,45", "4,武田信玄,95,60" });
        WriteXlsx(wideX, new string[] { "1,织田信长,100,55", "2,丰臣秀吉,90,40", "3,真田幸村,88,58", "4,武田信玄,95,60", "5,伊达政宗,70,40" });

        var t = Path.Combine(dir, "v2_x_base_apply.xlsx");
        File.Copy(baseX, t, true);
        var r = new ExcelDiffApplier().Apply(DoDiff(t, "Sheet1", wideX, "Sheet1"), t, "Sheet1", ExcelDiffDirection.DstToSrc);
        Check("xlsx dst→src 成功", r.IsSuccess && r.InsertedRowCount == 1, r.Message);

        Console.WriteLine("      指标(后) " + Metrics(wideX, "Sheet1", t, "Sheet1"));
        Check("xlsx 应用后与 wide 等价", Metrics(wideX, "Sheet1", t, "Sheet1") == "mod=0 add=0 rem=0 addRow=0 remRow=0", Metrics(wideX, "Sheet1", t, "Sheet1"));

        // 类型检查：第2行atk(55) 数值；插入的第5行(0基 index 4) id=5 数值
        var wb = NPOI.SS.UserModel.WorkbookFactory.Create(t);
        var sh = wb.GetSheet("Sheet1");
        var cellAtk = sh.GetRow(1).GetCell(3);
        Check("既有数字格保持 Numeric", cellAtk.CellType == NPOI.SS.UserModel.CellType.Numeric, "type=" + cellAtk.CellType);
        var cellNew = sh.GetRow(4).GetCell(0);
        Check("插入行数字格为 Numeric", cellNew != null && cellNew.CellType == NPOI.SS.UserModel.CellType.Numeric, "type=" + (cellNew == null ? "null" : cellNew.CellType.ToString()));
        var cellNewName = sh.GetRow(4).GetCell(1);
        Check("插入行文本格为 String", cellNewName != null && cellNewName.CellType == NPOI.SS.UserModel.CellType.String && cellNewName.StringCellValue == "伊达政宗", "value=" + (cellNewName == null ? "null" : cellNewName.StringCellValue));
    }

    // ============ 场景8：格子级显式应用（选中格） ============
        static void CellExplicitApply(string dir)
        {
            Console.WriteLine("== [场景] 格子级应用：只把第2行atk的值(50)写进 b，其余不动 ==");
            var srcB = Path.Combine(dir, "b.csv");
            var t = Path.Combine(dir, "v2_cell_apply.csv");
            File.Copy(srcB, t, true);

            var changes = new List<CellChange>();
            changes.Add(new CellChange(1, 3, "50"));
            var r = new ExcelDiffApplier().ApplyExplicit(t, "v2_cell_apply.csv", changes);
            Check("格子应用成功", r.IsSuccess && r.AppliedCellCount == 1, r.Message);

            var lines = File.ReadAllLines(t, Encoding.UTF8);
            Check("第2行atk=50", lines[1] == "1,织田信长,100,50", "got [" + lines[1] + "]");
            Check("第4行未被波及", lines[3] == "3,真田幸村,88,58", "got [" + lines[3] + "]");
            Check("新增行5保留", lines[5] == "5,伊达政宗,70,40", "got [" + lines[5] + "]");
        }

        // ============ 场景9：行级应用（ApplyRows：只同步选中行） ============
        static void RowLevelApply(string dir)
        {
            Console.WriteLine("== [场景] 行级应用：只同步选中行(新增行插入 / 修改行值回写) ==");
            var srcA = Path.Combine(dir, "a.csv");
            var srcB = Path.Combine(dir, "b.csv");

            // (a) 只同步"新增行"
            var t = Path.Combine(dir, "v3_row_apply.csv");
            File.Copy(srcA, t, true);
            var diff = DoDiff(t, "v3_row_apply.csv", srcB, "b.csv");
            int addedIdx = -1;
            foreach (var rd in diff.Rows)
                if (rd.Value.IsAdded()) { addedIdx = rd.Key; break; }
            var rows = new List<int>(); rows.Add(addedIdx);
            var r = new ExcelDiffApplier().ApplyRows(diff, t, "v3_row_apply.csv", ExcelDiffDirection.DstToSrc, rows);
            Check("新增行同步成功(插1行)", r.IsSuccess && r.InsertedRowCount == 1, r.Message);
            var lines = File.ReadAllLines(t, Encoding.UTF8);
            Check("行5已插入", lines[5] == "5,伊达政宗,70,40", "got [" + lines[5] + "]");
            Check("行1 atk 未被动(仍50)", lines[1] == "1,织田信长,100,50", "got [" + lines[1] + "]");
            var m = Metrics(srcB, "b.csv", t, "v3_row_apply.csv");
            Check("重对比: 新增行已同步,其余值保持原样(行1 atk + 行3 共4格)", m == "mod=4 add=0 rem=0 addRow=0 remRow=0", m);

            // (b) 只同步"修改行"的值
            var t2 = Path.Combine(dir, "v3_row_apply2.csv");
            File.Copy(srcA, t2, true);
            var diff2 = DoDiff(t2, "v3_row_apply2.csv", srcB, "b.csv");
            int modIdx = -1;
            foreach (var rd in diff2.Rows)
                if (!rd.Value.IsAdded() && !rd.Value.IsRemoved() && rd.Value.IsModified()) { modIdx = rd.Key; break; }
            var rows2 = new List<int>(); rows2.Add(modIdx);
            var r2 = new ExcelDiffApplier().ApplyRows(diff2, t2, "v3_row_apply2.csv", ExcelDiffDirection.DstToSrc, rows2);
            Check("修改行同步成功", r2.IsSuccess && r2.AppliedCellCount >= 1, r2.Message);
            var lines2 = File.ReadAllLines(t2, Encoding.UTF8);
            Check("行1 atk 已同步为55", lines2[1] == "1,织田信长,100,55", "got [" + lines2[1] + "]");
            Check("行3 未被动", lines2[3] == "3,德川家康,80,45", "got [" + lines2[3] + "]");
        }

        // ============ 场景10：列级应用（ApplyColumns） ============
        static void ColumnLevelApply(string dir)
        {
            Console.WriteLine("== [场景] 列级应用：只同步选中列(值列 / 结构列) ==");
            var srcA = Path.Combine(dir, "a.csv");
            var srcB = Path.Combine(dir, "b.csv");
            var srcW = Path.Combine(dir, "wide.csv");

            // (a) 值列：只同步 atk 列(对齐索引3)，名字/hp 不动
            var t = Path.Combine(dir, "v4_col_atk.csv");
            File.Copy(srcA, t, true);
            var diff = DoDiff(t, "v4_col_atk.csv", srcB, "b.csv");
            var r = new ExcelDiffApplier().ApplyColumns(diff, t, "v4_col_atk.csv", ExcelDiffDirection.DstToSrc, new List<int> { 3 });
            Check("atk 列值同步成功", r.IsSuccess && r.AppliedCellCount >= 2, r.Message);
            var lines = File.ReadAllLines(t, Encoding.UTF8);
            Check("行1 atk=55", lines[1] == "1,织田信长,100,55", "got [" + lines[1] + "]");
            Check("行3 仅 atk 变(名字/hp 未动)", lines[3] == "3,德川家康,80,58", "got [" + lines[3] + "]");
            Check("行2 未被动", lines[2] == "2,丰臣秀吉,90,40", "got [" + lines[2] + "]");

            // (b) 结构列：把 note 列整个插进 a
            var t2 = Path.Combine(dir, "v4_col_note.csv");
            File.Copy(srcA, t2, true);
            var diff2 = DoDiff(t2, "v4_col_note.csv", srcW, "wide.csv");
            var noteIdx = 4;   // id,name,hp,atk,note → 对齐索引4
            var r2 = new ExcelDiffApplier().ApplyColumns(diff2, t2, "v4_col_note.csv", ExcelDiffDirection.DstToSrc, new List<int> { noteIdx });
            Check("note 列插入成功", r2.IsSuccess && r2.InsertedColumnCount == 1, r2.Message);
            var lines2 = File.ReadAllLines(t2, Encoding.UTF8);
            Check("表头含 note", lines2[0] == "id,name,hp,atk,note", "got [" + lines2[0] + "]");
            Check("行1 note=n1", lines2[1] == "1,织田信长,100,50,n1", "got [" + lines2[1] + "]");
            Check("行4 note=n4", lines2[4] == "4,武田信玄,95,60,n4", "got [" + lines2[4] + "]");
        }

        static void WriteXlsx(string path, string[] lines)
    {
        var wb = new XSSFWorkbook();
        var sh = wb.CreateSheet("Sheet1");
        int r = 0;
        foreach (var line in lines)
        {
            var row = sh.CreateRow(r++);
            string[] fields = line.Split(',');
            for (int c = 0; c < fields.Length; c++)
            {
                var cell = row.CreateCell(c);
                double d;
                if (double.TryParse(fields[c], out d))
                    cell.SetCellValue(d);
                else
                    cell.SetCellValue(fields[c]);
            }
        }
        using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            wb.Write(fs);
    }

    static int Main(string[] args)
    {
        string dir = args[0];
        EnsureSample(dir);
        CsvRowSync(dir);
        CsvRowDelete(dir);
        CsvColumnSync(dir);
        RowLevelApply(dir);
        ColumnLevelApply(dir);
        CellExplicitApply(dir);
        XlsxRowAndType(dir);
        Console.WriteLine();
        if (fails == 0) { Console.WriteLine("ALL PASS"); return 0; }
        Console.WriteLine("FAILS=" + fails);
        return fails;
    }

    // 确保样例文件存在（每次运行都重写，保证测试确定性）
    static void EnsureSample(string dir)
    {
        var utf8 = new UTF8Encoding(false);
        File.WriteAllLines(Path.Combine(dir, "a.csv"), new string[] {
            "id,name,hp,atk", "1,织田信长,100,50", "2,丰臣秀吉,90,40", "3,德川家康,80,45", "4,武田信玄,95,60" }, utf8);
        File.WriteAllLines(Path.Combine(dir, "b.csv"), new string[] {
            "id,name,hp,atk", "1,织田信长,100,55", "2,丰臣秀吉,90,40", "3,真田幸村,88,58", "4,武田信玄,95,60", "5,伊达政宗,70,40" }, utf8);
        File.WriteAllLines(Path.Combine(dir, "c.csv"), new string[] {
            "id,name,hp,atk", "1,织田信长,100,50", "3,德川家康,80,45", "4,武田信玄,95,60" }, utf8);
        File.WriteAllLines(Path.Combine(dir, "wide.csv"), new string[] {
            "id,name,hp,atk,note", "1,织田信长,100,50,n1", "2,丰臣秀吉,90,40,n2", "3,德川家康,80,45,n3", "4,武田信玄,95,60,n4" }, utf8);
    }
}