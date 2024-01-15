using ClosedXML.Excel;
using Excely.ClosedXML.TableConverters;
using Excely.Workflows;

namespace Excely.Workflows
{
    public static class XlsxExporterExtension
    {
        /// <summary>
        /// 將指定的物件集合匯出至特定的 Excel 工作表。
        /// </summary>
        /// <param name="sourceData">來源資料</param>
        /// <param name="worksheet">指定的工作表</param>
        public static void ToWorksheet<TInput>(
            this ExcelyExporter<TInput> exporter, TInput sourceData,
            IXLWorksheet worksheet, CellLocation startCell = default)
        {
            var table = exporter.GetTable(sourceData);
            var tableWriter = new XlsxTableConverter(worksheet, startCell);
            tableWriter.ConvertFrom(table);
            foreach (var shaders in exporter.Shaders)
            {
                worksheet = shaders.Excute(worksheet);
            }
        }

        /// <summary>
        /// 將指定的物件集合匯出為全新的 Excel 實體。
        /// </summary>
        /// <param name="sourceData">來源資料</param>
        /// <param name="worksheetName">工作表名稱</param>
        /// <returns>全新的 Excel 實體</returns>
        public static XLWorkbook ToExcel<TInput>(
            this ExcelyExporter<TInput> exporter, TInput sourceData,
            string worksheetName = "sheet1",
            CellLocation startCell = default)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(worksheetName);
            exporter.ToWorksheet(sourceData, worksheet, startCell);
            return workbook;
        }
    }
}
