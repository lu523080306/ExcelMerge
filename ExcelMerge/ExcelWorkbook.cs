using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelMerge
{
    public class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; }

        public ExcelWorkbook()
        {
            Sheets = new Dictionary<string, ExcelSheet>();
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            if (Path.GetExtension(path) == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path) == ".tsv")
                return CreateFromTsv(path, config);

            var srcWb = WorkbookFactory.Create(path);
            var wb = new ExcelWorkbook();
            try
            {
                for (int i = 0; i < srcWb.NumberOfSheets; i++)
                {
                    var srcSheet = srcWb.GetSheetAt(i);
                    var sheet = ExcelSheet.Create(srcSheet, config);
                    sheet.Name = srcSheet.SheetName;
                    wb.Sheets.Add(srcSheet.SheetName, sheet);
                }
            }
            finally
            {
                // 释放 NPOI 底层文件句柄，否则对同一文件"读后再写"(apply)会报文件被占用
                ReleaseFileHandle(srcWb);
            }

            return wb;
        }

        public static IEnumerable<string> GetSheetNames(string path)
        {
            if (Path.GetExtension(path) == ".csv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else if (Path.GetExtension(path) == ".tsv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else
            {
                var names = new List<string>();
                var wb = WorkbookFactory.Create(path);
                try
                {
                    for (int i = 0; i < wb.NumberOfSheets; i++)
                        names.Add(wb.GetSheetAt(i).SheetName);
                }
                finally
                {
                    ReleaseFileHandle(wb);
                }

                foreach (var name in names)
                    yield return name;
            }
        }

        private static void ReleaseFileHandle(IWorkbook wb)
        {
            var xssf = wb as XSSFWorkbook;
            if (xssf == null)
                return;

            try { xssf.Package.Close(); } catch { }
        }

        private static ExcelWorkbook CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            var sheet = ExcelSheet.CreateFromCsv(path, config);
            sheet.Name = Path.GetFileName(path);
            wb.Sheets.Add(Path.GetFileName(path), sheet);

            return wb;
        }

        private static ExcelWorkbook CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            var sheet = ExcelSheet.CreateFromTsv(path, config);
            sheet.Name = Path.GetFileName(path);
            wb.Sheets.Add(Path.GetFileName(path), sheet);

            return wb;
        }
    }
}
