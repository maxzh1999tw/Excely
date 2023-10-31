using OfficeOpenXml;

namespace Excely.EPPlus.LGPL.Shaders.Xlsx
{
    /// <summary>
    /// 使儲存格自動適配寬高
    /// </summary>
    public class CellFittingShader : XlsxShaderBase
    {
        protected override void ExcuteOnWorksheet(ExcelWorksheet worksheet)
        {
            worksheet.Cells.AutoFitColumns();
            worksheet.Cells.Style.WrapText = true;
        }
    }
}
