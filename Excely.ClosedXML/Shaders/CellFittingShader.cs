using ClosedXML.Excel;

namespace Excely.ClosedXML.Shaders
{
    /// <summary>
    /// 使儲存格自動適配寬高。
    /// </summary>
    public class CellFittingShader : XlsxShaderBase
    {
        public double MinWidth { get; set; } = 1;
        public double MaxWidth { get; set; } = 50;

        protected override void ExcuteOnWorksheet(IXLWorksheet worksheet)
        {
            worksheet.Columns().AdjustToContents(MinWidth, MaxWidth);
            worksheet.Columns().Style.Alignment.WrapText = true;
        }
    }
}
