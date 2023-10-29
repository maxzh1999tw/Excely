using System.Reflection;

namespace Excely.TableFactorys
{
    /// <summary>
    /// 提供以類別結構為欄位，將物件集合傾印至表格的功能
    /// </summary>
    /// <typeparam name="T">欲轉換之類別</typeparam>
    public class ClassListTableFactory<T> : ITableFactory<IEnumerable<T>> where T : class
    {
        /// <summary>
        /// 決定 Property 是否應作為欄位匯出的執行邏輯，
        /// 輸入參數為 PropertyInfo，輸出結果為「是否應作為欄位匯出」，
        /// 若此屬性為 null，則全部欄位都匯出
        /// </summary>
        public Func<PropertyInfo, bool>? PropertyShowPolicy { get; set; }

        /// <summary>
        /// 決定 Property 作為欄位時的名稱
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位名稱」，
        /// 若此屬性為 null，則欄位名稱為 PropertyInfo.Name
        /// </summary>
        public Func<PropertyInfo, string?>? PropertyNamePolicy { get; set; }

        /// <summary>
        /// 決定 Property 作為欄位時的順序
        /// 輸入參數為 PropertyInfo，輸出結果為「排序(由小到大)」，
        /// 若此屬性為 null，則欄位依類別內預設排序
        /// </summary>
        public Func<PropertyInfo, int>? PropertyOrderPolicy { get; set; }

        /// <summary>
        /// 決定資料寫入欄位時的值
        /// 輸入參數為 (當前匯出物件, PropertyInfo)，輸出結果為「欲寫入欄位的值」，
        /// 若此屬性為 null，則值為該 Property 之 Value
        /// </summary>
        public Func<T, PropertyInfo, object?>? CustomValuePolicy { get; set; }

        public ExcelyTable GetTable(IEnumerable<T> sourceData)
        {
            var properties = typeof(T).GetProperties()
                                       .Where(x => PropertyShowPolicy?.Invoke(x) ?? true)
                                       .OrderBy(x => PropertyOrderPolicy?.Invoke(x) ?? 0)
                                       .ToList();

            var table = new List<IEnumerable<object?>>
            {
                GetSchema(properties)
            };

            foreach (var item in sourceData)
            {
                table.Add(GetRow(item, properties));
            }

            return new ExcelyTable(table);
        }

        /// <summary>
        /// 將類別結構轉換為表頭
        /// </summary>
        /// <param name="properties">欲匯出的 Properties</param>
        /// <returns>表頭列</returns>
        private IEnumerable<object?> GetSchema(IEnumerable<PropertyInfo> properties)
        {
            return properties.Select(property => PropertyNamePolicy?.Invoke(property) ?? property.Name).ToList();
        }

        /// <summary>
        /// 將物件轉換為資料列
        /// </summary>
        /// <param name="item">來源物件</param>
        /// <param name="properties">欲匯出的 Properties</param>
        /// <returns>資料列</returns>
        private IEnumerable<object?> GetRow(T item, IEnumerable<PropertyInfo> properties)
        {
            return properties.Select(property => CustomValuePolicy?.Invoke(item, property) ?? property.GetValue(item)).ToList();
        }
    }
}
