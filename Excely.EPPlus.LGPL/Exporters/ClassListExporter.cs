using Excely.EPPlus.LGPL.TableWriters;
using Excely.Shaders;
using Excely.TableFactories;
using OfficeOpenXml;
using System.Reflection;

namespace Excely.EPPlus.LGPL.Exporters
{
    /// <summary>
    /// 提供以類別結構為欄位，將物件集合匯出的功能
    /// </summary>
    /// <typeparam name="T">欲匯出的類別</typeparam>
    public class ClassListExporter<T> where T : class
    {
        /// <summary>
        /// 將 List 轉成 Table 的物件
        /// </summary>
        protected ClassListTableFactory<T> TableFactory { get; set; }

        /// <summary>
        /// 匯出資料完畢後依序執行的 IShader 集合
        /// </summary>
        public IEnumerable<IShader> Shaders { get; set; } = Enumerable.Empty<IShader>();

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="propertyShowPolicy">
        /// 決定 Property 是否應作為欄位匯出的執行邏輯，
        /// 輸入參數為 PropertyInfo，輸出結果為「是否應作為欄位匯出」，
        /// 若此屬性為 null，則全部欄位都匯出
        /// </param>
        /// <param name="propertyNamePolicy">
        /// 決定 Property 作為欄位時的名稱
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位名稱」，
        /// 若此屬性為 null，則欄位名稱為 PropertyInfo.Name
        /// </param>
        /// <param name="propertyOrderPolicy">
        /// 決定 Property 作為欄位時的順序
        /// 輸入參數為 PropertyInfo，輸出結果為「排序(由小到大)」，
        /// 若此屬性為 null，則欄位依類別內預設排序
        /// </param>
        /// <param name="customValuePolicy">
        /// 決定資料寫入欄位時的值
        /// 輸入參數為 (當前匯出物件, PropertyInfo)，輸出結果為「欲寫入欄位的值」，
        /// 若此屬性為 null，則值為該 Property 之 Value
        /// </param>
        public ClassListExporter(
            Func<PropertyInfo, bool>? propertyShowPolicy = null,
            Func<PropertyInfo, string?>? propertyNamePolicy = null,
            Func<PropertyInfo, int>? propertyOrderPolicy = null,
            Func<T, PropertyInfo, object?>? customValuePolicy = null
            )
        {
            TableFactory = new ClassListTableFactory<T>()
            {
                PropertyShowPolicy = propertyShowPolicy,
                PropertyNamePolicy = propertyNamePolicy,
                PropertyOrderPolicy = propertyOrderPolicy,
                CustomValuePolicy = customValuePolicy
            };
        }

        /// <summary>
        /// 將指定的物件集合匯出至特定的 Excel 工作表
        /// </summary>
        /// <param name="sourceData">來源資料</param>
        /// <param name="worksheet">指定的工作表</param>
        public void ToWorksheet(IEnumerable<T> sourceData, ExcelWorksheet worksheet)
        {
            var table = TableFactory.GetTable(sourceData);
            var tableWriter = new XlsxTableWriter(worksheet);
            tableWriter.Write(table);
            foreach (var shaders in Shaders)
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
        public ExcelPackage ToExcel(IEnumerable<T> sourceData, string worksheetName = "sheet1")
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(worksheetName);
            ToWorksheet(sourceData, worksheet);
            return package;
        }
    }
}
