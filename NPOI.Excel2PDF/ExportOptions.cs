using QuestPDF.Infrastructure;

namespace NPOI.Excel2PDF
{
    public class ExportOptions
    {
        public bool SeparateFilesPerSheet { get; set; } = false;
        public bool IncludeSheetNameInHeader { get; set; } = false;
        public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;
        public ContentDirection ContentDirection { get; set; }
        public float Margin { get; set; } = 20;
        public bool FitToPage { get; set; } = true;
        public float ScaleFactor { get; set; } = 1.0f;
        public bool IncludePageNumbers { get; set; } = true;
        public bool CompressPdf { get; set; } = true;
        public ImageCompressionQuality ImageCompressionQuality { get; set; }
        public bool PdfA { get; set; } = true;
    }
}
