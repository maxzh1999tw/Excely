using System.ComponentModel;
using System.Reflection;

namespace Excely.TableConverters
{
    /// <summary>
    /// 將 Table 轉換為物件列表。
    /// </summary>
    /// <typeparam name="TClass">目標類別</typeparam>
    public class ClassListTableConverter<TClass> : ITableConverter<IEnumerable<TClass>> where TClass : class, new()
    {
        /// <summary>
        /// 轉換過程的執行細節。
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
                // Lazy load
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
        public IEnumerable<TClass> ConvertFrom(ExcelyTable table)
        {
            if (Options.HasSchema)
            {
                // 有 Schema 就靠解析 Schema 來決定欄位
                return ImportInternal(table, (property, colIndex) =>
                {
                    // 取得 Property 轉換成的欄位名稱並比對
                    var name = Options.PropertyNamePolicy(property);
                    return name == table.Data[0][colIndex]?.ToString();
                });
            }
            else
            {
                // 沒有 Schema 依欄位順序解析
                return ImportInternal(table, (property, colIndex) =>
                {
                    // 取得 Property 應在的位置並比對
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

            // 遍歷每一 Row
            for (var rowIndex = Options.HasSchema ? 1 : 0; rowIndex < table.MaxRowCount; rowIndex++)
            {
                var obj = new TClass(); // 建立一個新目標物件
                var rowParseSuccess = true; // 此 Row 是否解析成功

                // 遍歷每一 Column 來取得欄位值
                for (var columnIndex = 0; columnIndex < table.MaxColumnCount; columnIndex++)
                {
                    // 取得這個 Column 代表的 Property
                    var property = TProperties.FirstOrDefault(p => propertyMatcher(p, columnIndex));
                    if (property == null)
                    {
                        // 這個 Column 沒有對應到任何 Property，不處理
                        continue;
                    }

                    // 取得欄位解析結果
                    var value = Options.PropertyValueSettingPolicy(property, table.Data[rowIndex][columnIndex]);

                    try
                    {
                        // 嘗試自動轉型
                        if (value != null && Options.EnableAutoTypeConversion)
                        {
                            var valueType = value.GetType();
                            if (!property.PropertyType.IsAssignableFrom(valueType))
                            {
                                // 嘗試以 TypeConverter 轉型
                                var typeConverter = TypeDescriptor.GetConverter(property.PropertyType);
                                if (typeConverter != null && typeConverter.CanConvertFrom(valueType))
                                {
                                    value = typeConverter.ConvertFrom(value);
                                }
                                else // 嘗試強制轉型
                                {
                                    var convertedValue = Convert.ChangeType(value, property.PropertyType);
                                    if (convertedValue != null)
                                    {
                                        value = convertedValue;
                                    }
                                }
                            }
                        }

                        property.SetValue(obj, value);
                    }
                    catch (Exception ex)
                    {
                        // 執行錯誤處理
                        var errorFixed = Options.ErrorHandlingPolicy(new CellLocation(rowIndex, columnIndex), obj, property, value, ex);

                        // 錯誤處理失敗，且要求停止
                        if (!errorFixed && Options.StopWhenError)
                        {
                            throw;
                        }

                        rowParseSuccess = errorFixed;
                    }

                    // 若解析失敗，放棄整 Row 的解析
                    if (!rowParseSuccess) break;
                }

                // 解析成功，加入結果
                if (rowParseSuccess) result.Add(obj);
            }

            return result;
        }
    }

    /// <summary>
    /// 定義一組 ClassListTableConverter 的執行細節。
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
        public Func<PropertyInfo, object?, object?> PropertyValueSettingPolicy { get; set; } = (prop, obj) => obj;

        /// <summary>
        /// 將值輸入進物件發生錯誤時，決定錯誤處理方式。
        /// 輸入參數為 (儲存格座標, 目標物件, PropertyInfo, 嘗試輸入的值, 發生的錯誤)，輸出結果為「錯誤是否得到修正」。
        /// 預設為不處理錯誤。
        /// </summary>
        public Func<CellLocation, TClass, PropertyInfo, object?, Exception, bool> ErrorHandlingPolicy { get; set; } = (_, _, _, _, _) => false;

        /// <summary>
        /// 當寫入的值與目標型別不同時，是否自動嘗試轉換。
        /// </summary>
        public bool EnableAutoTypeConversion { get; set; } = true;
    }
}
