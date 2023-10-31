using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Excely.Shaders.Xlsx
{
    /// <summary>
    /// 將一段範圍內的儲存格視為表格，並為其套上表格造型
    /// </summary>
    public class TableThemeShader : XlsxShaderBase
    {
        /// <summary>
        /// 表格左上角的儲存格座標
        /// </summary>
        public (int Row, int Col) StartCell { get; set; } = (0, 0);

        /// <summary>
        /// 表頭高度
        /// </summary>
        public int SchemaHeight { get; set; } = 1;

        /// <summary>
        /// 表格寬度
        /// </summary>
        public int TableWidth { get; set; }

        /// <summary>
        /// 表格高度
        /// </summary>
        public int TableHeight { get; set; }

        /// <summary>
        /// 表格主題
        /// </summary>
        public TableTheme Theme { get; set; } = TableTheme.Default;


        public TableThemeShader(int tableWidth, int tableHeight)
        {
            TableWidth = tableWidth;
            TableHeight = tableHeight;
        }

        protected override void ExcuteOnWorksheet(ExcelWorksheet worksheet)
        {
            // 設定標題的背景色和文字色
            var header = worksheet.Cells[StartCell.Row + 1, StartCell.Col + 1, StartCell.Row + SchemaHeight, StartCell.Col + TableWidth];
            header.Style.Fill.PatternType = ExcelFillStyle.Solid;
            header.Style.Fill.BackgroundColor.SetColor(Theme.HeaderBackgroundColor);
            header.Style.Font.Color.SetColor(Theme.HeaderTextColor);
            header.Style.Font.Bold = true;

            // 給整個表格設定邊框色
            var table = worksheet.Cells[StartCell.Row + 1, StartCell.Col + 1, StartCell.Row + TableHeight, StartCell.Col + TableWidth];
            table.Style.Border.Top.Style = table.Style.Border.Bottom.Style = table.Style.Border.Left.Style = table.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            table.Style.Border.Top.Color.SetColor(Theme.BorderColor);
            table.Style.Border.Bottom.Color.SetColor(Theme.BorderColor);
            table.Style.Border.Left.Color.SetColor(Theme.BorderColor);
            table.Style.Border.Right.Color.SetColor(Theme.BorderColor);
        }
    }

    /// <summary>
    /// 用於表達表格主題的結構
    /// </summary>
    public struct TableTheme
    {
        /// <summary>
        /// 表頭背景色
        /// </summary>
        public Color HeaderBackgroundColor { get; set; }

        /// <summary>
        /// 表頭文字顏色
        /// </summary>
        public Color HeaderTextColor { get; set; }

        /// <summary>
        /// 邊框顏色
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
