using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace TheFirstProject
{
    [System.Runtime.InteropServices.GuidAttribute("B488E567-3BE0-40C1-B1DF-4922A7FCAF21")]
    class Program
    {
        static void Main(string[] args)
        {
            GenerateExcelFromList2();
        }

        static void GenerateExcelFromList2()
        {
            var fileName = "GenerateExcelFromList.xlsx";

            ExcelPackage excelPackage = CreateWorkbook( fileName);
            using (excelPackage)
            {
                var worksheet = CreateSheet(excelPackage, "GenerateExcelFromList");
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Name-------------";
                worksheet.Cells[1, 2].Style.Font.Bold = true;

                FormatCell(worksheet, 1, 1, 1, 5);

                excelPackage.Save();
            }
        }

        private static ExcelPackage CreateWorkbook(string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            FileInfo file = new FileInfo(fileName);
            ExcelPackage excelPackage = new ExcelPackage(file);
            return excelPackage;
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

        //Format the rows
        private static void FormatCell(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            using (var range = worksheet.Cells[fromRow, fromCol, toRow, toCol])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.Blue);
                range.Style.Font.Color.SetColor(Color.WhiteSmoke);
                range.Style.ShrinkToFit = false;
            }

            var cell = worksheet.Cells["C1"];
            cell.Hyperlink = new Uri("http://www.google.com");
            cell.Value = "Click me!";
        }
    }
}
