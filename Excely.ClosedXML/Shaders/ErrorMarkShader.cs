using ClosedXML.Excel;

namespace Excely.ClosedXML.Shaders
{
    /// <summary>
    /// 將出錯的表格醒目標示，並標註錯誤原因。
    /// </summary>
    public class ErrorMarkShader : XlsxShaderBase
    {
        /// <summary>
        /// 匯入錯誤資料。
        /// </summary>
        public Dictionary<CellLocation, string> CellErrors { get; set; }

        /// <summary>
        /// 醒目標示文字顏色。
        /// </summary>
        public XLColor TextColor { get; set; } = XLColor.Red;

        /// <param name="errorCells">匯入錯誤資料</param>
        /// <param name="author">Comment 作者(不得為空)</param>
        public ErrorMarkShader(Dictionary<CellLocation, string> errorCells)
        {
            CellErrors = errorCells;
        }

        protected override void ExcuteOnWorksheet(IXLWorksheet worksheet)
        {
            foreach (var cellError in CellErrors)
            {
                var cell = worksheet.Cell(cellError.Key.Row + 1, cellError.Key.Column + 1);
                cell.Style.Font.SetFontColor(TextColor);
                cell.GetComment().AddText(cellError.Value);
            }
        }
    }
}
