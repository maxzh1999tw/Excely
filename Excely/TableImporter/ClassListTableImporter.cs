using System.Reflection;

namespace Excely.TableImporter
{
    /// <summary>
    /// 提供將 Table 匯入為 Class list 的功能
    /// </summary>
    /// <typeparam name="TClass">欲轉換的類別</typeparam>
    public class ClassListTableImporter<TClass> : ITableImporter<IEnumerable<TClass>> where TClass : class, new()
    {
        /// <summary>
        /// 匯入的 Table 是否含有表頭。
        /// 當此欄位為 true 時，會使用 PropertyNamePolicy；
        /// 否則會使用 PropertyIndexPolicy。
        /// </summary>
        public bool HasSchema { get; set; } = true;

        /// <summary>
        /// 當轉換發生錯誤時是否立刻停止。
        /// </summary>
        public bool StopWhenError { get; set; } = true;

        /// <summary>
        /// 決定 Property 作為欄位時的名稱。
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位名稱」，
        /// 若此屬性為 null，則欄位名稱為 PropertyInfo.Name。
        /// </summary>
        public Func<PropertyInfo, string?>? PropertyNamePolicy { get; set; }

        /// <summary>
        /// 決定 Property 出現在表頭時的位置。
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位索引」，
        /// 若該 Property 沒有出現在表頭中，請回傳 null。
        /// </summary>
        public Func<PropertyInfo, int?>? PropertyIndexPolicy { get; set; }

        /// <summary>
        /// 決定將值寫入至 Property 時應寫入的值。
        /// 輸入參數為 (PropertyInfo, 原始值)，輸出結果為「應寫入的值」。
        /// </summary>
        public Func<PropertyInfo, object?, object?>? CustomValuePolicy { get; set; }

        private PropertyInfo[]? _TProperies;

        /// <summary>
        /// 目標型別 TClass 的 Properies。
        /// </summary>
        private PropertyInfo[] TProperties
        {
            get
            {
                _TProperies ??= typeof(TClass).GetProperties();
                return _TProperies;
            }
        }

        /// <summary>
        /// 將指定的 Table 匯入為 Class list。
        /// </summary>
        /// <returns>匯入結果</returns>
        public ImportResult<IEnumerable<TClass>> Import(ExcelyTable table)
        {
            if (HasSchema)
            {
                return ImportInternal(table, (property, colIndex) =>
                {
                    var name = PropertyNamePolicy == null ? property.Name : PropertyNamePolicy(property);
                    return name == table.Data[0][colIndex]?.ToString();
                });
            }
            else
            {
                return ImportInternal(table, (property, colIndex) =>
                {
                    var index = PropertyIndexPolicy == null ? Array.IndexOf(TProperties, property) : PropertyIndexPolicy(property);
                    return index == colIndex;
                });
            }
        }

        /// <summary>
        /// 讀取 Table 資料並轉換為 Object 的流程。
        /// </summary>
        /// <param name="table">來源 Table</param>
        /// <param name="propertyMatcher">取得欄位與 Property 對應的邏輯</param>
        /// <returns>匯入結果</returns>
        private ImportResult<IEnumerable<TClass>> ImportInternal(ExcelyTable table, Func<PropertyInfo, int, bool> propertyMatcher)
        {
            var result = new List<TClass>(table.MaxRowCount);
            var rowErrors = new Dictionary<CellLocation, Exception>();

            for (var row = HasSchema ? 1 : 0; row < table.MaxRowCount; row++)
            {
                var obj = new TClass();
                var rowParseSuccess = true;

                for (var col = 0; col < table.MaxColCount; col++)
                {
                    var property = TProperties.FirstOrDefault(p => propertyMatcher(p, col));

                    if (property != null)
                    {
                        var value = CustomValuePolicy != null
                            ? CustomValuePolicy(property, table.Data[row][col])
                            : table.Data[row][col];

                        try
                        {
                            property.SetValue(obj, value);
                        }
                        catch (Exception ex)
                        {
                            if (StopWhenError) throw;
                            rowErrors[new CellLocation(row, col)] = ex;
                            rowParseSuccess = false;
                        }
                    }
                }

                if (rowParseSuccess) result.Add(obj);
            }

            return new ImportResult<IEnumerable<TClass>>(result, rowErrors);
        }
    }
}
