using Excely.EPPlus.LGPL.TableWriters;
using OfficeOpenXml;

namespace Excely.Exporters
{
    /// <summary>
    /// 為 ClassListExporter 提供 Excel 相關的擴展
    /// </summary>
    public static class ClassListExporterExtension
    {
        /// <summary>
        /// 將指定的物件集合匯出至特定的 Excel 工作表
        /// </summary>
        /// <param name="sourceData">來源資料</param>
        /// <param name="worksheet">指定的工作表</param>
        public static void ToWorksheet<T>(this ClassListExporter<T> exporter, IEnumerable<T> sourceData, ExcelWorksheet worksheet) where T : class
        {
            var table = exporter.TableFactory.GetTable(sourceData);
            var tableWriter = new XlsxTableWriter(worksheet);
            tableWriter.Write(table);
            foreach (var shaders in exporter.Shaders)
            {
                worksheet = shaders.Excute(worksheet);
            }
        }

        /// <summary>
        /// 將指定的物件集合匯出為全新的 Excel 實體
        /// </summary>
        /// <param name="sourceData">來源資料</param>
        /// <param name="worksheetName">工作表名稱</param>
        /// <returns>全新的 Excel 實體</returns>
        public static ExcelPackage ToExcel<T>(this ClassListExporter<T> exporter, IEnumerable<T> sourceData, string worksheetName = "sheet1") where T : class
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(worksheetName);
            exporter.ToWorksheet(sourceData, worksheet);
            return package;
        }
    }
}
