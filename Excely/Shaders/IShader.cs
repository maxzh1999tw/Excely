using OfficeOpenXml;

namespace Excely.Shaders
{
    /// <summary>
    /// 定義了一個可以調整工作表的執行單元
    /// </summary>
    public interface IShader
    {
        /// <summary>
        /// 對特定工作表執行調整
        /// </summary>
        /// <param name="worksheet">目標工作表</param>
        void Excute(ExcelWorksheet worksheet);
    }
}
