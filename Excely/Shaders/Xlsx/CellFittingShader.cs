using OfficeOpenXml;

namespace Excely.Shaders.Xlsx
{
    /// <summary>
    /// 使儲存格自動適配寬高
    /// </summary>
    public class CellFittingShader : IShader
    {
        public void Excute(ExcelWorksheet worksheet)
        {
            worksheet.Cells.AutoFitColumns();
            worksheet.Cells.Style.WrapText = true;
        }
    }
}
