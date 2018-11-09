using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using ExcelObject = Microsoft.Office.Interop.Excel;

namespace ExcelToXML.Functions
{
    public class ExcelFile
    {
        public string FileName { get; set; }

        public ExcelObject.Application excelApp { set; get; } = null;
        public ExcelObject.Workbooks books { set; get; } = null;
        public ExcelObject.Workbook sheet { set; get; } = null;
        public ExcelObject.Worksheet worksheet { set; get; } = null;
    }

    public static class ExcelFunctions
    {
        public static ExcelFile OpenExcelFile(string file)
        {
            ExcelFile result = new ExcelFile() { FileName = file };
            try
            {
                result.excelApp = new ExcelObject.Application();
                result.excelApp.Visible = false;
                result.books = result.excelApp.Workbooks;
                result.sheet = result.books.Open(result.FileName);
            }
            catch
            {
                CloseExcelFile(result);
                result = null;
            }

            return result;
        }

        public static void CloseExcelFile(ExcelFile excelFile)
        {
            if (excelFile?.worksheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFile.worksheet);
            }
            excelFile?.sheet?.Close(false);
            excelFile?.books?.Close();
            if (excelFile?.sheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFile.sheet);
            }
            if (excelFile?.books != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFile.books);
            }
            excelFile?.excelApp?.Quit();
            if (excelFile?.excelApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFile?.excelApp);
            }
        }

        public static string GetString(ExcelObject.Worksheet worksheet, int row, int col)
        {
            string result = "";

            ExcelObject.Range range = worksheet.Cells[row, col];
            try
            {
                result = range.Value2?.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            }

            return result;
        }
    }
}
