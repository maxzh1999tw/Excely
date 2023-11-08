using System.Formats.Asn1;
using System.Reflection;

namespace Excely.TableFactories
{
    /// <summary>
    /// 提供以類別結構為欄位，將物件集合傾印至表格的功能。
    /// </summary>
    /// <typeparam name="TClass">欲轉換之類別</typeparam>
    public class ClassListTableFactory<TClass> : ITableFactory<IEnumerable<TClass>> where TClass : class
    {
        /// <summary>
        /// 轉換過程的執行細節
        /// </summary>
        protected ClassListTableFactoryOptions<TClass> Options { get; set; } = new ClassListTableFactoryOptions<TClass>();

        #region === 建構子 ===
        public ClassListTableFactory() { }

        public ClassListTableFactory(ClassListTableFactoryOptions<TClass> options)
        {
            Options = options;
        }
        #endregion

        public ExcelyTable GetTable(IEnumerable<TClass> sourceData)
        {
            var properties = typeof(TClass).GetProperties();
            properties = properties
                            .Where(x => Options.PropertyShowPolicy(x))
                            .OrderBy(x => Options.PropertyOrderPolicy(properties, x))
                            .ToArray();

            var table = new List<IList<object?>>
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
        /// 將類別結構轉換為表頭。
        /// </summary>
        /// <param name="properties">欲匯出的 Properties</param>
        /// <returns>表頭列</returns>
        private IList<object?> GetSchema(IEnumerable<PropertyInfo> properties)
        {
            return properties.Select(property => Options.PropertyNamePolicy(property)).ToList<object?>();
        }

        /// <summary>
        /// 將物件轉換為資料列。
        /// </summary>
        /// <param name="item">來源物件</param>
        /// <param name="properties">欲匯出的 Properties</param>
        /// <returns>資料列</returns>
        private IList<object?> GetRow(TClass item, IEnumerable<PropertyInfo> properties)
        {
            return properties.Select(property => Options.CustomValuePolicy(property, item)).ToList();
        }
    }

    /// <summary>
    /// 定義一組 ClassListTableFactory 的執行細節
    /// </summary>
    /// <typeparam name="TClass">目標類別</typeparam>
    public class ClassListTableFactoryOptions<TClass>
    {
        /// <summary>
        /// 決定 Property 是否應作為欄位匯出的執行邏輯。
        /// 輸入參數為 PropertyInfo，輸出結果為「是否應作為欄位匯出」，
        /// 預設為全部欄位都匯出。
        /// </summary>
        public Func<PropertyInfo, bool> PropertyShowPolicy { get; set; } = _ => true;

        /// <summary>
        /// 決定 Property 作為欄位時的名稱。
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位名稱」，
        /// 預設為 PropertyInfo.Name。
        /// </summary>
        public Func<PropertyInfo, string?> PropertyNamePolicy { get; set; } = p => p.Name;

        /// <summary>
        /// 決定 Property 作為欄位時的順序。
        /// 輸入參數為 (所有Property, 當前PropertyInfo)，輸出結果為「排序(由小到大)」，
        /// 預設為依類別內預設排序。
        /// </summary>
        public Func<PropertyInfo[], PropertyInfo, int> PropertyOrderPolicy { get; set; } = (propertys, p) => Array.IndexOf(propertys, p);

        /// <summary>
        /// 決定資料寫入欄位時的值。
        /// 輸入參數為 (PropertyInfo, 當前匯出物件)，輸出結果為「欲寫入欄位的值」，
        /// 預設為該 Property 之 Value。
        /// </summary>
        public Func<PropertyInfo, TClass, object?> CustomValuePolicy { get; set; } = (p, obj) => p.GetValue(obj);
    }
}
