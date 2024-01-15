using ClosedXML.Excel;
using System.Drawing;

namespace Excely.ClosedXML.Shaders
{
    /// <summary>
    /// 將一段範圍內的儲存格視為表格，並為其套上表格造型。
    /// </summary>
    public class TableThemeShader : XlsxShaderBase
    {
        /// <summary>
        /// 表格左上角的儲存格座標。
        /// </summary>
        public CellLocation StartCell { get; set; } = new(0, 0);

        /// <summary>
        /// 表頭高度。
        /// </summary>
        public int SchemaHeight { get; set; } = 1;

        /// <summary>
        /// 表格寬度。
        /// </summary>
        public int TableWidth { get; set; }

        /// <summary>
        /// 表格高度。
        /// </summary>
        public int TableHeight { get; set; }

        /// <summary>
        /// 表格主題。
        /// </summary>
        public TableTheme Theme { get; set; } = TableTheme.Default;

        public TableThemeShader() { }

        /// <param name="tableWidth">表格寬度(0為自適應)</param>
        /// <param name="tableHeight">表格高度(0為自適應)</param>
        public TableThemeShader(int tableWidth, int tableHeight)
        {
            TableWidth = tableWidth;
            TableHeight = tableHeight;
        }

        protected override void ExcuteOnWorksheet(IXLWorksheet worksheet)
        {
            var tableWidth = TableWidth;
            if (TableWidth == 0)
            {
                tableWidth = worksheet.LastColumnUsed().ColumnNumber();
            }

            var tableHeight = TableHeight;
            if (TableHeight == 0)
            {
                tableHeight = worksheet.LastRowUsed().RowNumber();
            }

            // 設定標題的背景色和文字色
            var header = worksheet.Range(StartCell.Row + 1, StartCell.Column + 1, StartCell.Row + SchemaHeight, StartCell.Column + tableWidth);
            header.Style.Fill.BackgroundColor = XLColor.FromColor(Theme.HeaderBackgroundColor);
            header.Style.Font.FontColor = XLColor.FromColor(Theme.HeaderTextColor);
            header.Style.Font.Bold = true;

            // 給整個表格設定邊框色
            var table = worksheet.Range(StartCell.Row + 1, StartCell.Column + 1, StartCell.Row + tableHeight, StartCell.Column + tableWidth);
            table.Style.Border.OutsideBorder = table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            table.Style.Border.OutsideBorderColor = table.Style.Border.InsideBorderColor = XLColor.FromColor(Theme.BorderColor);
        }
    }

    /// <summary>
    /// 用於表達表格主題的結構。
    /// </summary>
    public struct TableTheme
    {
        /// <summary>
        /// 表頭背景色。
        /// </summary>
        public Color HeaderBackgroundColor { get; set; }

        /// <summary>
        /// 表頭文字顏色。
        /// </summary>
        public Color HeaderTextColor { get; set; }

        /// <summary>
        /// 邊框顏色。
        /// </summary>
        public Color BorderColor { get; set; }

        public static TableTheme Default => new()
        {
            HeaderBackgroundColor = Color.FromArgb(68, 114, 196),
            HeaderTextColor = Color.White,
            BorderColor = Color.FromArgb(68, 114, 196),
        };
    }
}