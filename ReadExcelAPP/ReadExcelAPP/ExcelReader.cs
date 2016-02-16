namespace ReadExcelAPP
{
    using System;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    internal interface IExcelReader : IDisposable
    {
        CellReadResponse<decimal> ReadCellAsDecimal(int sheetIndex, int rowIndex, int columnIndex);

        CellReadResponse<DateTime> ReadCellAsDateTime(int sheetIndex, int rowIndex, int columnIndex);

        CellReadResponse<string> ReadCellAsString(int sheetIndex, int rowIndex, int columnIndex);
    }

    internal class ExcelReader : IExcelReader
    {
        private readonly MemoryStream documentMemoryStream;
        private readonly SpreadsheetDocument document;
        private readonly ICellReferenceBuilder cellReferenceBuilder;

        public ExcelReader(MemoryStream documentMemoryStream, SpreadsheetDocument document, ICellReferenceBuilder cellReferenceBuilder)
        {
            this.documentMemoryStream = documentMemoryStream;
            this.document = document;
            this.cellReferenceBuilder = cellReferenceBuilder;
        }

        public CellReadResponse<decimal> ReadCellAsDecimal(int sheetIndex, int rowIndex, int columnIndex)
        {
            var result = new CellReadResponse<decimal> { Status = CellReadStatus.Success };

            if (document.WorkbookPart.WorksheetParts.Count() <= sheetIndex)
            {
                result.Status = CellReadStatus.EmptyCell;
            }
            else
            {
                var sheetData = document.WorkbookPart.WorksheetParts.ElementAt(sheetIndex).Worksheet.Elements<SheetData>().Single();
                var cell = sheetData.Descendants<Cell>().SingleOrDefault(c => c.CellReference == cellReferenceBuilder.GetCellReference(rowIndex, columnIndex));
                if (cell == null || cell.CellValue == null || string.IsNullOrEmpty(cell.CellValue.Text))
                {
                    result.Status = CellReadStatus.EmptyCell;
                }
                else if (cell.DataType != null && cell.DataType != CellValues.Number)
                {
                    result.Status = CellReadStatus.WrongDataType;
                }
                else
                {
                    result.Value = (decimal)Convert.ToDouble(cell.CellValue.Text);
                }
            }

            return result;
        }

        public CellReadResponse<DateTime> ReadCellAsDateTime(int sheetIndex, int rowIndex, int columnIndex)
        {
            var result = new CellReadResponse<DateTime> { Status = CellReadStatus.Success };

            if (document.WorkbookPart.WorksheetParts.Count() <= sheetIndex)
            {
                result.Status = CellReadStatus.EmptyCell;
            }
            else
            {
                var sheetData = document.WorkbookPart.WorksheetParts.ElementAt(sheetIndex).Worksheet.Elements<SheetData>().Single();
                var cell = sheetData.Descendants<Cell>().SingleOrDefault(c => c.CellReference == cellReferenceBuilder.GetCellReference(rowIndex, columnIndex));
                if (cell == null || cell.CellValue == null || string.IsNullOrEmpty(cell.CellValue.Text))
                {
                    result.Status = CellReadStatus.EmptyCell;
                }
                else if (cell.DataType != null && cell.DataType != CellValues.Date)
                {
                    result.Status = CellReadStatus.WrongDataType;
                }
                else
                {
                    var excelDate = Convert.ToDouble(cell.CellValue.Text);
                    result.Value = DateTime.FromOADate(excelDate);
                }
            }

            return result;
        }

        public CellReadResponse<string> ReadCellAsString(int sheetIndex, int rowIndex, int columnIndex)
        {
            var result = new CellReadResponse<string> { Status = CellReadStatus.Success };

            if (document.WorkbookPart.WorksheetParts.Count() <= sheetIndex)
            {
                result.Status = CellReadStatus.EmptyCell;
            }
            else
            {
                var sheetData = document.WorkbookPart.WorksheetParts.ElementAt(sheetIndex).Worksheet.Elements<SheetData>().Single();
                var cell = sheetData.Descendants<Cell>().SingleOrDefault(c => c.CellReference == cellReferenceBuilder.GetCellReference(rowIndex, columnIndex));
                if (cell == null || cell.CellValue == null || string.IsNullOrEmpty(cell.CellValue.Text))
                {
                    result.Status = CellReadStatus.EmptyCell;
                }
                else if (cell.DataType == null || (cell.DataType != CellValues.String && cell.DataType != CellValues.SharedString))
                {
                    result.Status = CellReadStatus.WrongDataType;
                }
                else
                {
                    if (cell.DataType == CellValues.SharedString)
                    {
                        var sharedStringIndex = Convert.ToInt32(cell.CellValue.Text);
                        var sharedString = document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
                        result.Value = sharedString.Text.Text;
                    }
                    else
                    {
                        result.Value = cell.CellValue.Text;
                    }
                }
            }

            return result;
        }

        public void Dispose()
        {
            document.Dispose();
            documentMemoryStream.Dispose();
        }
    }
}
