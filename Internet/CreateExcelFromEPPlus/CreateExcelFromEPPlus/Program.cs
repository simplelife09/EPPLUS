using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace CreateExcelFromEPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            GenerateExcelFromList2();
        }

        static void GenerateExcelFromList()
        {
            var dataRepo = new DataRepository();
            var bookList = dataRepo.GetBookList();
            var fileName = "GenerateExcelFromList.xlsx";
            if (File.Exists("GenerateExcelFromList.xlsx"))
            {
                File.Delete("GenerateExcelFromList.xlsx");
            }
            var file = new FileInfo(fileName);
            using (var excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("GenerateExcelFromList");
                worksheet.Cells["A1"].LoadFromCollection(bookList);
                var worksheet2 = excelPackage.Workbook.Worksheets.Add("GenerateExcelFromList2");
                worksheet2.Cells["A1"].LoadFromCollection(bookList);
                excelPackage.Save();
            }            
        }

        static void GenerateExcelFromList2()
        {
            var dataRepo = new DataRepository();
            var bookList = dataRepo.GetBookList();
            var bookDataTable = dataRepo.GetBookDataTable();
            var fileName = "GenerateExcelFromList.xlsx";
            if (File.Exists("GenerateExcelFromList.xlsx"))
            {
                File.Delete("GenerateExcelFromList.xlsx");
            }
            var file = new FileInfo(fileName);
            using (var excelPackage = new ExcelPackage(file))
            {
                var worksheet = CreateSheet(excelPackage, "GenerateExcelFromList");
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 2].Style.Font.Bold = true;
                worksheet.Cells["A2"].LoadFromCollection(bookList);
                var worksheet2 = excelPackage.Workbook.Worksheets.Add("GenerateExcelFromList2");
                worksheet2.Cells["A1"].LoadFromDataTable(bookDataTable, true);
                excelPackage.Save();
            }
        }

        private static ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[1];
            ws.Name = sheetName; //Setting Sheet's name
            ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

            return ws;
        }
    }
}
