using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace NPOICore
{
    class Program
    {
        static void Main(string[] args)
        {
            var outputFileName = "PortfolioImportTemplate.xlsx";
            using (var templateStream = new FileStream(outputFileName, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook excel = UpdateExcel(templateStream, 5000, true);

                using (var outputStream = new FileStream(outputFileName, FileMode.Create, FileAccess.Write))
                {
                    excel.Write(outputStream);
                }
            }
        }

        private static XSSFWorkbook UpdateExcel(FileStream templateStream, int length, bool createCodeColumns)
        {
            var excel = new XSSFWorkbook(templateStream);
            var sheet = (XSSFSheet)excel.GetSheetAt(0);

            CreateMetaDataRows(excel.CreateSheet("Area Manager"), DataSources.AreaManagers, createCodeColumns);
            SetCodeColumn(sheet, "F", 6, length, "Area Manager", createCodeColumns);
            CreateMetaDataRows(excel.CreateSheet("Sales"), DataSources.Sales, createCodeColumns);
            SetCodeColumn(sheet, "G", 7, length, "Sales", createCodeColumns);
            CreateMetaDataRows(excel.CreateSheet("Products"), DataSources.Products, createCodeColumns);
            SetCodeColumn(sheet, "H", 8, length, "Products", createCodeColumns);
            return excel;
        }

        private static void SetCodeColumn(XSSFSheet sheet, string columnLetter, int colindex, int elements, string sheetName, bool createCodeColumns)
        {
            if (createCodeColumns)
            {
                var row0 = sheet.GetRow(0);
                var lastColIndex = row0.LastCellNum;
                var headerCell = row0.CreateCell(lastColIndex, CellType.String);
                headerCell.SetCellValue(sheetName + " Code");
                row0.Cells.Add(headerCell);
                headerCell.CellStyle = sheet.Workbook?.GetSheetAt(0)?.GetRow(0)?.GetCell(0)?.CellStyle;

                for (var i = 1; i <= elements; i++)
                {
                    var row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    var cell = row.CreateCell(lastColIndex, CellType.Formula);
                    var formula = $"INDEX('{sheetName}'!A1:D{elements},MATCH(${columnLetter}{i + 1},'{sheetName}'!D1:D{elements},0),1)";
                    cell.SetCellFormula(formula);
                }
                sheet.SetColumnHidden(row0.LastCellNum - 1, true);
            }

            var validationHelper = new XSSFDataValidationHelper(sheet);
            var addressList = new CellRangeAddressList(0, elements - 1, colindex - 1, colindex - 1);
            var constraint = validationHelper.CreateFormulaListConstraint($"'{sheetName}'!$D$2:$D$" + elements);
            var dataValidation = validationHelper.CreateValidation(constraint, addressList);
            sheet.AddValidationData(dataValidation);
        }

        private static void CreateMetaDataRows(ISheet sheet, List<DataSource> items, bool createCodeColumns)
        {
            CreateRow(sheet, "Code", "First Name", "Last Name", "Full Name", 0);
            for (var i = 1; i < items.Count; i++)
            {
                var fullName = createCodeColumns ? $"{items[i].LastName} {items[i].FirstName}" : $"{items[i].LastName} {items[i].FirstName} ({items[i].Code})";
                CreateRow(sheet, items[i].Code, items[i].FirstName, items[i].LastName, fullName, i);
            }
            sheet.SetColumnHidden(3, true);
            Enumerable.Range(0, 4).ToList().ForEach(i => sheet.AutoSizeColumn(i));
        }

        private static void CreateRow(ISheet sheet, string code, string firstname, string lastname, string fullname, int i)
        {
            IRow row = sheet.CreateRow(i);

            var codeCell = row.CreateCell(0);
            codeCell.SetCellValue(code);

            var fnCell = row.CreateCell(1);
            fnCell.SetCellValue(firstname);

            var lnCell = row.CreateCell(2);
            lnCell.SetCellValue(lastname);

            var funCell = row.CreateCell(3);
            funCell.SetCellValue(fullname);

            if (i == 0)
            {
                var styleCell = sheet.Workbook.GetSheetAt(0).GetRow(0).GetCell(0);
                codeCell.CellStyle = styleCell.CellStyle;
                fnCell.CellStyle = styleCell.CellStyle;
                lnCell.CellStyle = styleCell.CellStyle;
                funCell.CellStyle = styleCell.CellStyle;
            }
        }
    }
}
