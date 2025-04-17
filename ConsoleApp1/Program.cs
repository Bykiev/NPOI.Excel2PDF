using NPOI.Excel2PDF;

namespace TestApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var options = new ExportOptions
            {
                SeparateFilesPerSheet = false,
                IncludeSheetNameInHeader = true,
                Orientation = PageOrientation.Landscape,
                FitToPage = false,
                IncludePageNumbers = true,
                CompressPdf = true
            };

            var res = ExcelToPdfConverter.Convert("rotate.xls", options);

            foreach (var item in res)
                File.WriteAllBytes($"output{item.SheetName}.pdf", item.Data);
        }
    }
}
