namespace ReadExcelAPP
{
    using System.IO;

    class Program
    {
        
        static void Main(string[] args)
        {
            IExcelProcessorFactory excelProcessorFactory;
            IExcelReader excelReader;

            var excelFile = File.ReadAllBytes(@"D:\Data\temp.xlsx");
            ICellReferenceBuilder cellReferenceBuilder = new CellReferenceBuilder();
            excelProcessorFactory = new ExcelProcessorFactory(cellReferenceBuilder);
            using (excelReader = excelProcessorFactory.CreateExcelReader(excelFile))
            {
                var readNameResponse = excelReader.ReadCellAsString(0, 2, 0);
                if (readNameResponse.Status == CellReadStatus.Success)
                {
                    var name = readNameResponse.Value;
                }

            }

        }
    }
}
