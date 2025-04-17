using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using QuestPDF;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;

namespace NPOI.Excel2PDF
{
    public static class ExcelToPdfConverter
    {
        private static IFormulaEvaluator evaluator;
        private static readonly DataFormatter dataFormatter;

        static ExcelToPdfConverter()
        {
            Settings.License = LicenseType.Community;
            dataFormatter = new DataFormatter();
        }

        public static List<ExportResult> Convert(string inputPath, ExportOptions options)
        {
            IWorkbook workbook = null;

            try
            {
                using (FileStream file = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                {
                    if (Path.GetExtension(inputPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                        workbook = new XSSFWorkbook(file);
                    else
                        workbook = new HSSFWorkbook(file);
                }

                return Convert(workbook, options);
            }
            finally
            {
                workbook?.Dispose();
            }
        }

        public static List<ExportResult> Convert(IWorkbook workbook, ExportOptions options)
        {
            var files = new List<ExportResult>();
            evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

            if (options.SeparateFilesPerSheet)
            {
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var data = ConvertSingleSheet(workbook, i, options);
                    string sheetName = workbook.GetSheetName(i);

                    files.Add(new ExportResult()
                    {
                        SheetName = sheetName,
                        Data = data
                    });
                }
            }
            else
            {
                files.Add(new ExportResult()
                {
                    SheetName = "",
                    Data = ConvertAllSheets(workbook, options)
                });
            }

            return files;
        }

        private static byte[] ConvertAllSheets(IWorkbook workbook, ExportOptions options)
        {
            string author = "";
            string title = "";
            string subject = "";
            string keywords = "";

            if (workbook is XSSFWorkbook)
            {
                var metadata = ((XSSFWorkbook)workbook).GetProperties();

                author = metadata.CoreProperties.Creator;
                title = metadata.CoreProperties.Title;
                subject = metadata.CoreProperties.Subject;
                keywords = metadata.CoreProperties.Keywords;
            }
            else
            {
                var metadata = ((HSSFWorkbook)workbook).SummaryInformation;

                if (metadata != null)
                {
                    author = metadata.Author;
                    title = metadata.Title;
                    subject = metadata.Subject;
                    keywords = metadata.Keywords;
                }
            }

            return Document.Create(container =>
            {
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    container.Page(page =>
                    {
                        ConfigurePage(page, options, sheet.SheetName);

                        page.Content()
                            .Column(column =>
                            {
                                if (options.IncludeSheetNameInHeader)
                                {
                                    column.Item()
                                        .PaddingBottom(10)
                                        .Text(sheet.SheetName)
                                        .FontSize(16)
                                        .Bold();
                                }

                                ConvertSheetToPdf(workbook, sheet, column, options);
                            });
                    });
                }
            })
            .WithMetadata(new DocumentMetadata()
            {
                Author = author,
                Title = title,
                Subject = subject,
                Keywords = keywords,
                CreationDate = DateTimeOffset.Now,
                ModifiedDate = DateTimeOffset.Now,
                Creator = "NPOI.Excel2PDF",
                Producer = "NPOI.Excel2PDF"
            })
            .GeneratePdf();
        }

        private static byte[] ConvertSingleSheet(IWorkbook workbook, int sheetIndex, ExportOptions options)
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);

            return Document.Create(container =>
            {
                container.Page(page =>
                {
                    ConfigurePage(page, options, sheet.SheetName);

                    page.Content()
                        .Column(column =>
                        {
                            if (options.IncludeSheetNameInHeader)
                            {
                                column.Item()
                                    .PaddingBottom(10)
                                    .Text(sheet.SheetName)
                                    .FontSize(16)
                                    .Bold();
                            }

                            ConvertSheetToPdf(workbook, sheet, column, options);
                        });
                })
                ;
            })
            .WithSettings(new DocumentSettings()
            {
                CompressDocument = options.CompressPdf,
                PdfA = options.PdfA,
                ContentDirection = options.ContentDirection,
                ImageCompressionQuality = options.ImageCompressionQuality
            })
            .GeneratePdf();
        }

        private static void ConfigurePage(PageDescriptor page, ExportOptions options, string sheetName)
        {
            page.Size(options.Orientation == PageOrientation.Portrait ?
                PageSizes.A4 : PageSizes.A4.Landscape());
            page.MarginVertical(options.Margin);
            page.MarginHorizontal(options.Margin);

            if (options.IncludePageNumbers)
            {
                page.Footer().AlignCenter().Text(text =>
                {
                    text.CurrentPageNumber();
                    text.Span(" / ");
                    text.TotalPages();
                });
            }

            page.Header()
                .Text(sheetName)
                .FontSize(12)
                .SemiBold();
        }

        private static void ConvertSheetToPdf(IWorkbook workbook, ISheet sheet, ColumnDescriptor column, ExportOptions options)
        {
            int maxColumnCount = sheet.GetMaxColumnCount();

            // calculate scale
            float scale = CalculateScale(sheet, options);

            // create table in PDF
            column.Item().Scale(scale).Table(table =>
            {
                // define columns
                table.ColumnsDefinition(_ =>
                {
                    for (int colIdx = 0; colIdx <= maxColumnCount; colIdx++)
                    {
                        if (!sheet.IsColumnHidden(colIdx))
                        {
                            var width = (float)sheet.GetColumnWidthInPixels(colIdx);
                            _.RelativeColumn(width);
                        }
                    }
                });

                // add rows
                for (int rowIdx = sheet.FirstRowNum; rowIdx <= sheet.LastRowNum; rowIdx++)
                {
                    IRow row = sheet.GetRow(rowIdx);
                    if (row != null && row.Hidden != true)
                    {
                        for (int colIdx = 0; colIdx <= maxColumnCount; colIdx++)
                        {
                            if (!sheet.IsColumnHidden(colIdx))
                            {
                                var excelCell = row.GetCell(colIdx);
                                if (excelCell != null)
                                {
                                    uint currentRow = (uint)(rowIdx + 1);
                                    uint currentCol = (uint)(colIdx + 1);

                                    if (IsMergedCell(sheet, rowIdx, colIdx))
                                    {
                                        var region = GetMergedRegion(sheet, rowIdx, colIdx);
                                        if (region.FirstRow == rowIdx && region.FirstColumn == colIdx)
                                        {
                                            int columnCount = region.LastColumn - region.FirstColumn + 1;

                                            table.Cell()
                                             .Row(currentRow)
                                             .Column(currentCol)
                                             .ColumnSpan((uint)columnCount)
                                             .RowSpan((uint)(region.LastRow - region.FirstRow + 1))
                                             .Element(ProcessCell(workbook, excelCell));
                                        }
                                    }
                                    else
                                    {
                                        table.Cell()
                                            .Row(currentRow)
                                            .Column(currentCol)
                                            .Element(ProcessCell(workbook, excelCell));
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }

        private static Action<IContainer> ProcessCell(IWorkbook workbook, ICell cell)
        {
            return container =>
            {
                float columnWidth = (float)SheetUtil.GetColumnWidth(cell.Sheet, cell.ColumnIndex, true);
                if (columnWidth == -1)
                    columnWidth = 8;

                ICellStyle style = cell.CellStyle;

                if (style.Rotation != 0)
                {
                    float angle = style.Rotation;

                    if (workbook is HSSFWorkbook)
                    {
                        if (angle >= 0 && angle <= 90)
                            angle = 360 - angle;
                        else
                            angle = -angle;
                    }
                    else
                    {
                        if (angle >= 0 && angle <= 90)
                            angle = 360 - angle;
                        else
                            angle -= 90;
                    }

                    container = container.Rotate(angle);

                    // TODO: we should call TranslateY(val) here to move content down after rotation
                }

                container = container.MinHeight(cell.Row.HeightInPoints);
                container = container.MinWidth(columnWidth);

                //switch (style.VerticalAlignment)
                //{
                //    case SS.UserModel.VerticalAlignment.Top:
                //        container = container.AlignTop();
                //        break;

                //    case SS.UserModel.VerticalAlignment.Bottom:
                //        container = container.AlignBottom();
                //        break;

                //    case SS.UserModel.VerticalAlignment.Center:
                //        container = container.AlignMiddle();
                //        break;
                //}

                container
                    .BorderTop(GetBorderSize(style.BorderTop))
                    .BorderRight(GetBorderSize(style.BorderRight))
                    .BorderBottom(GetBorderSize(style.BorderBottom))
                    .BorderLeft(GetBorderSize(style.BorderLeft))
                    .Background(GetBackgroundColor(style))
                    //.BorderColor(Colors.Grey.Lighten1)
                    .Text(text =>
                    {
                        switch (style.Alignment)
                        {
                            case SS.UserModel.HorizontalAlignment.Center:
                                text.AlignCenter();
                                break;

                            case SS.UserModel.HorizontalAlignment.Left:
                                text.AlignLeft();
                                break;

                            case SS.UserModel.HorizontalAlignment.Right:
                                text.AlignRight();
                                break;

                            case SS.UserModel.HorizontalAlignment.Justify:
                                text.Justify();
                                break;
                        }

                        string cellText = GetFormattedCellValue(cell);
                        IFont font = cell.Sheet.Workbook.GetFontAt(style.FontIndex);

                        if (cell.Hyperlink == null)
                        {
                            var span = text.Span(cellText)
                                .FontSize((float)font.FontHeightInPoints)
                                .FontFamily(font.FontName);

                            if (workbook is XSSFWorkbook)
                                span.FontColor(GetXssfFontColor(font));
                            else
                                span.FontColor(GetHssfFontColor(font, (HSSFWorkbook)workbook));

                            if (font.TypeOffset == FontSuperScript.Super)
                                span.Superscript();
                            else if (font.TypeOffset == FontSuperScript.Sub)
                                span.Subscript();

                            if (font.IsBold)
                                span.Bold();

                            span.Italic(font.IsItalic);
                            span.Underline(font.Underline != FontUnderlineType.None);
                            span.Strikethrough(font.IsStrikeout);
                        }
                        else
                        {
                            var hyperlinkStyle = TextStyle.Default
                                .FontColor(Colors.Blue.Medium)
                                .FontSize((float)font.FontHeightInPoints)
                                .FontFamily(font.FontName)
                                .Underline();

                            if (font.IsBold)
                                hyperlinkStyle.Bold();

                            hyperlinkStyle.Italic(font.IsItalic);
                            hyperlinkStyle.Underline(font.Underline != FontUnderlineType.None);
                            hyperlinkStyle.Strikethrough(font.IsStrikeout);

                            hyperlinkStyle.Italic(font.IsItalic);

                            text.Hyperlink(cellText, cell.Hyperlink.Address)
                                .Style(hyperlinkStyle);
                        }
                    });
            };
        }

        private static string GetFormattedCellValue(ICell cell)
        {
            if (cell.CellType == CellType.Numeric)
            {
                if (DateUtil.IsCellDateFormatted(cell))
                    return cell.DateCellValue.ToString();
            }

            return dataFormatter.FormatCellValue(cell, evaluator);
        }

        private static Color GetBackgroundColor(ICellStyle style)
        {

            if (style.FillForegroundColorColor != null)
            {
                if (style.FillForegroundColorColor is HSSFColor)
                {
                    if (style.FillForegroundColor == HSSFColor.Automatic.Index)
                        return Colors.White;
                }

                var color = style.FillForegroundColorColor.RGB;
                if (color != null)
                    return Color.FromRGB(color[0], color[1], color[2]);
            }

            return Colors.White;
        }

        private static string GetXssfFontColor(IFont font)
        {
            if (font is XSSFFont)
            {
                var xssfFont = (XSSFFont)font;
                var color = xssfFont.GetXSSFColor();

                if (color != null)
                    return color.ARGBHex;
            }

            return "#000000";
        }

        private static Color GetHssfFontColor(IFont font, HSSFWorkbook workbook)
        {
            if (font is HSSFFont)
            {
                var hssfFont = (HSSFFont)font;
                var color = hssfFont.GetHSSFColor(workbook);

                if (color != null)
                    return Color.FromRGB(color.RGB[0], color.RGB[1], color.RGB[2]);
            }

            return Colors.Black;
        }

        private static bool IsMergedCell(ISheet sheet, int row, int col)
        {
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress region = sheet.GetMergedRegion(i);
                if (region.FirstRow <= row && region.LastRow >= row &&
                    region.FirstColumn <= col && region.LastColumn >= col)
                {
                    return true;
                }
            }
            return false;
        }

        private static CellRangeAddress GetMergedRegion(ISheet sheet, int row, int col)
        {
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress region = sheet.GetMergedRegion(i);
                if (region.FirstRow <= row && region.LastRow >= row &&
                    region.FirstColumn <= col && region.LastColumn >= col)
                {
                    return region;
                }
            }
            return null;
        }

        private static float CalculateScale(ISheet sheet, ExportOptions options)
        {
            if (!options.FitToPage)
                return options.ScaleFactor;

            // get the PDF page size in pixels
            float pageWidth = options.Orientation == PageOrientation.Portrait
                ? PageSizes.A4.Width - options.Margin * 2
                : PageSizes.A4.Height - options.Margin * 2;

            // get the Excel table size in pixels
            double tableWidth = 0;
            int lastColumn = sheet.GetMaxColumnCount();
            for (int colIdx = 0; colIdx <= lastColumn; colIdx++)
                tableWidth += sheet.GetColumnWidthInPixels(colIdx);

            float widthScale = pageWidth / (float)tableWidth;
            float finalScale = Math.Min(widthScale, 1.0f);

            float minScale = 0.3f; // minimum scale is 30%
            finalScale = Math.Max(finalScale, minScale);

            // apply user-defined scale
            if (options.ScaleFactor != 1.0f)
                finalScale = Math.Min(finalScale, options.ScaleFactor);

            return finalScale;
        }

        private static float GetBorderSize(BorderStyle style)
        {
            switch (style)
            {
                case BorderStyle.None:
                    return 0;
                case BorderStyle.Thick:
                    return 1.2f;
                case BorderStyle.Medium:
                case BorderStyle.MediumDashDot:
                case BorderStyle.MediumDashDotDot:
                    return 0.7f;
                default:
                    return 0.3f;
            }
        }
    }
}
