using Excely.Shaders;
using Excely.TableFactorys;
using Excely.TableWriters;
using OfficeOpenXml;

namespace Excely.Exporters
{
    /// <summary>
    /// 提供以類別結構為欄位，將物件集合匯出的功能
    /// </summary>
    /// <typeparam name="T">欲匯出的類別</typeparam>
    public class ClassListExporter<T> : ClassListTableFactory<T> where T : class
    {
        /// <summary>
        /// 匯出資料完畢後依序執行的 IShader 集合
        /// </summary>
        public IEnumerable<IShader> Shaders { get; set; } = Enumerable.Empty<IShader>();

        /// <summary>
        /// 將指定的物件集合匯出至特定的 Excel 工作表
        /// </summary>
        /// <param name="sourceData">來源資料</param>
        /// <param name="worksheet">指定的工作表</param>
        public void ToWorksheet(IEnumerable<T> sourceData, ExcelWorksheet worksheet)
        {
            var table = GetTable(sourceData);
            var tableWriter = new XlsxTableWriter(worksheet);
            tableWriter.Write(table);
            foreach (var shaders in Shaders)
            {
                shaders.Excute(worksheet);
            }
        }

        /// <summary>
        /// 將指定的物件集合匯出為全新的 Excel 實體
        /// </summary>
        /// <param name="sourceData">來源資料</param>
        /// <param name="worksheetName">工作表名稱</param>
        /// <returns>全新的 Excel 實體</returns>
        public ExcelPackage ToExcel(IEnumerable<T> sourceData, string worksheetName = "sheet1")
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(worksheetName);
            ToWorksheet(sourceData, worksheet);
            return package;
        }
    }
}
