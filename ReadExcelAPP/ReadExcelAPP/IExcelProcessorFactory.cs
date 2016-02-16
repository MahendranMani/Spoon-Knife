
namespace ReadExcelAPP
{
    using System.IO;

    using DocumentFormat.OpenXml.Packaging;

    internal interface IExcelProcessorFactory
    {
        IExcelReader CreateExcelReader(byte[] fileBytes);
    }

    internal class ExcelProcessorFactory : IExcelProcessorFactory
    {
        protected readonly ICellReferenceBuilder CellReferenceBuilder;

        public ExcelProcessorFactory(ICellReferenceBuilder cellReferenceBuilder)
        {
            CellReferenceBuilder = cellReferenceBuilder;
        }

        public IExcelReader CreateExcelReader(byte[] fileBytes)
        {
            var documentMemoryStream = new MemoryStream(fileBytes);
            var document = SpreadsheetDocument.Open(documentMemoryStream, false);

            var reader = new ExcelReader(documentMemoryStream, document, CellReferenceBuilder);
            return reader;
        }

    }
}
