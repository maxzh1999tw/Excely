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
            if (sourceData == null) throw new ArgumentNullException(nameof(sourceData));

            // 取得要匯出的 Properties
            var properties = typeof(TClass).GetProperties();
            properties = properties
                            .Where(x => Options.PropertyShowPolicy(x))
                            .OrderBy(x => Options.PropertyOrderPolicy(properties, x))
                            .ToArray();

            var table = new List<IList<object?>>();

            // 加入表頭
            if (Options.WithSchema)
            {
                table.Add(GetSchema(properties));
            }

            // 加入資料
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
    /// 定義一組 ClassListTableFactory 的執行細節。
    /// </summary>
    /// <typeparam name="TClass">目標類別</typeparam>
    public class ClassListTableFactoryOptions<TClass>
    {
        /// <summary>
        /// 決定匯出時是否帶有表頭。
        /// 預設為是。
        /// </summary>
        public bool WithSchema { get; set; } = true;

        /// <summary>
        /// 決定 Property 是否應作為欄位匯出。
        /// 預設為全部欄位都匯出。
        /// </summary>
        public PropertyShowPolicyDelegate PropertyShowPolicy { get; set; } = _ => true;

        /// <summary>
        /// 決定 Property 作為欄位時的名稱。
        /// 預設為 PropertyInfo.Name。
        /// </summary>
        public PropertyNamePolicyDelegate PropertyNamePolicy { get; set; } = property => property.Name;

        /// <summary>
        /// 決定 Property 作為欄位時的順序。
        /// 預設為依類別內預設排序。
        /// </summary>
        public PropertyOrderPolicyDelegate PropertyOrderPolicy { get; set; } = (properties, property) => Array.IndexOf(properties, property);

        /// <summary>
        /// 決定資料寫入欄位時的值。
        /// 預設為該 Property 之 Value。
        /// </summary>
        public CustomValuePolicyDelegate CustomValuePolicy { get; set; } = (property, obj) => property.GetValue(obj);

        #region ===== Policy delegates =====

        /// <summary>
        /// 決定 Property 是否應作為欄位匯出。
        /// </summary>
        /// <param name="property">當前決定的 Property</param>
        /// <returns>是否應匯出</returns>
        public delegate bool PropertyShowPolicyDelegate(PropertyInfo property);

        /// <summary>
        /// 決定 Property 作為欄位時的名稱。
        /// </summary>
        /// <param name="property">當前決定的 Property</param>
        /// <returns>欄位名稱</returns>
        public delegate string? PropertyNamePolicyDelegate(PropertyInfo property);

        /// <summary>
        /// 決定 Property 作為欄位時的順序。
        /// </summary>
        /// <param name="allProperties">目標類別的所有 Property</param>
        /// <param name="property">當前決定的 Property</param>
        /// <returns>欄位權重(越小越靠前)</returns>
        public delegate int? PropertyOrderPolicyDelegate(PropertyInfo[] allProperties, PropertyInfo property);

        /// <summary>
        /// 決定資料寫入欄位時的值。
        /// </summary>
        /// <param name="property">當前決定的 Property</param>
        /// <param name="obj">當前寫入的目標物件</param>
        /// <returns>應寫入的值</returns>
        public delegate object? CustomValuePolicyDelegate(PropertyInfo property, TClass obj);

        #endregion ===== Policy delegates =====
    }
}
