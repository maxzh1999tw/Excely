using System.ComponentModel;
using System.Reflection;

namespace Excely.TableConverters
{
    /// <summary>
    /// 將 Table 轉換為物件列表
    /// </summary>
    /// <typeparam name="TClass">目標類別</typeparam>
    public class ClassListTableConverter<TClass> : ITableConverter<IEnumerable<TClass>> where TClass : class, new()
    {
        /// <summary>
        /// 轉換過程的執行細節
        /// </summary>
        protected ClassListTableConverterOptions<TClass> Options { get; set; } = new ClassListTableConverterOptions<TClass>();

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

        #region === 建構子 ===
        public ClassListTableConverter() { }

        public ClassListTableConverter(ClassListTableConverterOptions<TClass> options)
        {
            Options = options;
        }
        #endregion


        /// <summary>
        /// 將指定的 Table 轉換為 Class list。
        /// </summary>
        /// <returns>轉換結果</returns>
        public IEnumerable<TClass> Convert(ExcelyTable table)
        {
            if (Options.HasSchema)
            {
                return ImportInternal(table, (property, colIndex) =>
                {
                    var name = Options.PropertyNamePolicy(property);
                    return name == table.Data[0][colIndex]?.ToString();
                });
            }
            else
            {
                return ImportInternal(table, (property, colIndex) =>
                {
                    var index = Options.PropertyIndexPolicy(TProperties, property);
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
        private IEnumerable<TClass> ImportInternal(ExcelyTable table, Func<PropertyInfo, int, bool> propertyMatcher)
        {
            var result = new List<TClass>(table.MaxRowCount);
            for (var row = Options.HasSchema ? 1 : 0; row < table.MaxRowCount; row++)
            {
                var obj = new TClass();
                var rowParseSuccess = true;
                for (var col = 0; col < table.MaxColCount; col++)
                {
                    var property = TProperties.FirstOrDefault(p => propertyMatcher(p, col));
                    if (property != null)
                    {
                        var value = Options.CustomValuePolicy(property, table.Data[row][col]);
                        try
                        {
                            var valueType = value?.GetType();
                            if (Options.DoAutoConvert &&
                                value != null &&
                                !property.PropertyType.IsAssignableFrom(value.GetType()))
                            {
                                var typeConverter = TypeDescriptor.GetConverter(property.PropertyType);
                                if (valueType != null && typeConverter != null && typeConverter.CanConvertFrom(valueType))
                                {
                                    value = typeConverter.ConvertFrom(value);
                                }
                            }
                            property.SetValue(obj, value);
                        }
                        catch (Exception ex)
                        {
                            var ignore = Options.ErrorHandlingPolicy(new CellLocation(row, col), obj, property, value, ex);
                            if (!ignore && Options.StopWhenError)
                            {
                                throw;
                            }
                            rowParseSuccess = ignore;
                        }
                    }
                    if (!rowParseSuccess) break;
                }
                if (rowParseSuccess) result.Add(obj);
            }
            return result;
        }
    }

    /// <summary>
    /// 定義一組 ClassListTableConverter 的執行細節
    /// </summary>
    /// <typeparam name="TClass"></typeparam>
    public class ClassListTableConverterOptions<TClass>
    {
        /// <summary>
        /// 匯入的 Table 是否含有表頭。
        /// 當此欄位為 true 時，會使用 PropertyNamePolicy；
        /// 否則會使用 PropertyIndexPolicy。
        /// </summary>
        public bool HasSchema { get; set; } = true;

        /// <summary>
        /// 當轉換發生錯誤時是否立刻停止。
        /// 若此欄為 false，則發生錯誤時會跳過該 Row，繼續執行匯入。
        /// </summary>
        public bool StopWhenError { get; set; } = true;

        /// <summary>
        /// 決定 Property 作為欄位時的名稱。
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位名稱」，
        /// 預設為 Property.Name。
        /// </summary>
        public Func<PropertyInfo, string?> PropertyNamePolicy { get; set; } = prop => prop.Name;

        /// <summary>
        /// 決定 Property 出現在表頭時的位置。
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位索引」，
        /// 若該 Property 沒有出現在表頭中，請回傳 null。
        /// 預設為依照 property 預設順序排序。
        /// </summary>
        public Func<PropertyInfo[], PropertyInfo, int?> PropertyIndexPolicy { get; set; } = (props, prop) => Array.IndexOf(props, prop);

        /// <summary>
        /// 決定將值寫入至 Property 時應寫入的值。
        /// 輸入參數為 (PropertyInfo, 原始值)，輸出結果為「應寫入的值」。
        /// 預設為原值。
        /// </summary>
        public Func<PropertyInfo, object?, object?> CustomValuePolicy { get; set; } = (prop, obj) => obj;

        /// <summary>
        /// 將值輸入進物件發生錯誤時，決定錯誤處理方式。
        /// 輸入參數為 (儲存格座標, 目標物件, PropertyInfo, 嘗試輸入的值, 發生的錯誤, 是否忽略此錯誤)，
        /// 預設為不處理錯誤。
        /// </summary>
        public Func<CellLocation, TClass, PropertyInfo, object?, Exception, bool> ErrorHandlingPolicy { get; set; } = (_, _, _, _, _) => false;

        /// <summary>
        /// 當寫入的值與目標型別不同時，是否自動嘗試轉換。
        /// </summary>
        public bool DoAutoConvert { get; set; } = true;
    }
}
