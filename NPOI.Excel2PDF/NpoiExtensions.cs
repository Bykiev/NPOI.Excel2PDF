using NPOI.SS.UserModel;

namespace NPOI.Excel2PDF
{
    public static class NpoiExtensions
    {
        public static bool IsNullOrEmpty(this ICell cell)
        {
            if (cell != null)
            {
                //if (cell.CellStyle != null)
                //    return false;

                switch (cell.CellType)
                {
                    case CellType.String:
                        return string.IsNullOrWhiteSpace(cell.StringCellValue);
                    case CellType.Boolean:
                    case CellType.Numeric:
                    case CellType.Formula:
                    case CellType.Error:
                        return false;
                }
            }

            return true;
        }

        public static int GetMaxColumnCount(this ISheet sheet)
        {
            int lastCol = 0;
            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);

                if (row != null)
                {
                    int rowLastCol = 0;

                    for (int col = 0; col <= row.LastCellNum; col++)
                    {
                        var cell = row.GetCell(col);
                        if (!cell.IsNullOrEmpty())
                            rowLastCol = col;
                    }

                    if (rowLastCol > lastCol)
                        lastCol = rowLastCol;
                }
            }

            return lastCol;
        }
    }
}