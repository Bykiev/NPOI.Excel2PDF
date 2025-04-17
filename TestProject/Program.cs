using NPOI.Excel2PDF;

namespace TestProject
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var converter = new ExcelToPdfConverter();
            var options = new ExportOptions
            {
                SeparateFilesPerSheet = false,
                IncludeSheetNameInHeader = true,
                Orientation = PageOrientation.Landscape,
                FitToPage = true,
                IncludePageNumbers = true,
                Author = "My Company",
                Title = "Financial Report",
                CompressPdf = true
            };

            converter.ConvertExcelToPdf("input.xlsx", "output.pdf", options);
        }
    }
}
