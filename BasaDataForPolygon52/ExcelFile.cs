using System;
using System.IO;
using OfficeOpenXml;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace BasaDataForPolygon52
{
    public static class DBExcel
    {
        public static string locationFileExcel = Application.StartupPath + "/Bd01.xlsx";

        public static ExcelWorksheet ObjCurrentSheet = null;
        public static ExcelPackage Excel;

        public static int CurrentSheetEND = 0;

        public static void Initialize()
        {
            /*
            using (var pck = new ExcelPackage())
            {
                using (var stream = File.OpenRead(locationFileExcel))
                {
                    pck.Load(stream);
                    Excel = pck;
                }
            }
            */
            
            FileInfo existingFile = new FileInfo(locationFileExcel);
            ExcelPackage package = new ExcelPackage(existingFile);
            Excel = package;
        }

        public static bool CheckFile()
        {
            return Excel.File.Exists;
        }

        public static void InitializeSheet(ListBox listbox)
        {
            ObjCurrentSheet = Excel.Workbook.Worksheets[listbox.SelectedIndex + 2];
        }

        private static object getValue(int column, int row)
        {
            return Excel.Workbook.Worksheets[1].Cells[column, row].Value;
        }

        public static string getName(int column)
        {
            int Name = 3;
            if (Excel.Workbook.Worksheets[1].Cells[column, Name].Value == null)
                return "_НАЗВАНИЕ";
            return Excel.Workbook.Worksheets[1].Cells[column, Name].Value.ToString();
        }

        public static string getVendorCode(int column)
        {
            int VendorCode = 4;
            if (Excel.Workbook.Worksheets[1].Cells[column, VendorCode].Value == null)
                return "_АРТИКУЛ";
            return Excel.Workbook.Worksheets[1].Cells[column, VendorCode].Value.ToString();
        }

        public static int getCount(int column)
        {
            int Count = 6;
            if (Excel.Workbook.Worksheets[1].Cells[column, Count].Value == null)
                return 0;
            return Convert.ToInt32(Excel.Workbook.Worksheets[1].Cells[column, Count].Value);
        }

        public static int getPrice(int column)
        {
            int Price = 7;
            if (Excel.Workbook.Worksheets[1].Cells[column, Price].Value == null)
                return 0;
            return Convert.ToInt32(Excel.Workbook.Worksheets[1].Cells[column, Price].Value);
        }

        public static string getStock(int column)
        {
            int Stock = 11;
            if (Excel.Workbook.Worksheets[1].Cells[column, Stock].Value != null)
            {
                if (Excel.Workbook.Worksheets[1].Cells[column, Stock].Value.ToString() == "1")
                    return "Акция";
                else
                    return string.Empty;
            }
            else
                return string.Empty;
        }

        public static string getBrand(int column)
        {
            int Brand = 15;
            if (Excel.Workbook.Worksheets[1].Cells[column, Brand].Value == null)
                return "_БРЭНД";
            return Excel.Workbook.Worksheets[1].Cells[column, Brand].Value.ToString();
        }

        public static string getSection(int column)
        {
            int Section = 17;
            if (Excel.Workbook.Worksheets[1].Cells[column, Section].Value == null)
                return "_РАЗДЕЛ";
            return Excel.Workbook.Worksheets[1].Cells[column, Section].Value.ToString();
        }

        public static string getSubSection(int column)
        {
            int SubSection = 16;
            if (Excel.Workbook.Worksheets[1].Cells[column, SubSection].Value == null)
                return "_ПОДРАЗДЕЛ";
            return Excel.Workbook.Worksheets[1].Cells[column, SubSection].Value.ToString();
        }

        public static string getModelProduct(int column)
        {
            int ModelProduct = 18;
            if (Excel.Workbook.Worksheets[1].Cells[column, ModelProduct].Value == null)
                return "_МОДЕЛЬ ПРОДУКТА";
            return Excel.Workbook.Worksheets[1].Cells[column, ModelProduct].Value.ToString();
        }

        public static void newCountForPosition(int position, int newCount)
        {
            int Count = 6;
            int oldCount = Convert.ToInt32(Excel.Workbook.Worksheets[1].Cells[position, Count].Value);
            Excel.Workbook.Worksheets[1].Cells[position, Count].Value = oldCount - newCount;
        }

        public static void SaveExcelFile()
        {
            Excel.Save();
        }




        //Возвращает номер последней строки на листе;
        public static int EndDB()
        {
            //ExcelRangeBase[] spisok = new ExcelRangeBase[0];
            int i;
            bool endSheet = true;

            for (i = 1; endSheet; i++)
            {
                if (Excel.Workbook.Worksheets[1].Cells[i, 1].Value != null)
                    if (Excel.Workbook.Worksheets[1].Cells[i, 1].Value.ToString() == "КОМИССИЯ")
                        break;
            }
            return i;
        }
    }
}
