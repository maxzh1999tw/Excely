﻿namespace Excely.TableConverters
{
    /// <summary>
    /// 將 Table 轉換為字典列表。
    /// </summary>
    public class DictionaryListTableConverter : ITableConverter<IEnumerable<Dictionary<string, object?>>>
    {
        /// <summary>
        /// 轉換過程的執行細節。
        /// </summary>
        protected DictionaryListTableConverterOptions Options { get; set; } = new DictionaryListTableConverterOptions();

        #region === 建構子 ===
        public DictionaryListTableConverter() { }

        public DictionaryListTableConverter(DictionaryListTableConverterOptions options)
        {
            Options = options;
        }
        #endregion

        /// <summary>
        /// 將指定的 Table 轉換為 Dictionary list。
        /// </summary>
        /// <returns>轉換結果</returns>
        public IEnumerable<Dictionary<string, object?>> ConvertFrom(ExcelyTable table)
        {
            // 取得 key 集合
            var keys = new string?[table.MaxColumnCount];
            if (Options.HasSchema)
            {
                // 有表頭時，會傳入欄位名稱
                var schema = table.Data[0];
                for (int i = 0; i < keys.Length; i++)
                {
                    keys[i] = Options.CustomKeyNamePolicy(i, schema[i]?.ToString());
                }
            }
            else
            {
                // 沒表頭時只傳欄位 index
                for (int i = 0; i < keys.Length; i++)
                {
                    keys[i] = Options.CustomKeyNamePolicy(i, null);
                }
            }

            var result = new List<Dictionary<string, object?>>();

            // 遍歷每一 Row
            for (var rowIndex = Options.HasSchema ? 1 : 0; rowIndex < table.MaxRowCount; rowIndex++)
            {
                // 建立新 Dictionary
                var dict = new Dictionary<string, object?>();

                // 遍歷每一 Column
                for (var columnIndex = 0; columnIndex < table.MaxColumnCount; columnIndex++)
                {
                    // 取得欄位對應的 Key
                    var key = keys[columnIndex];

                    // 若為 null 代表不轉換此欄位
                    if (key == null) continue;

                    dict.Add(key, Options.CustomValuePolicy(key, table.Data[rowIndex][columnIndex]));
                }

                result.Add(dict);
            }

            return result;
        }
    }

    /// <summary>
    /// 定義一組 DictionaryListTableConverter 的執行細節。
    /// </summary>
    public class DictionaryListTableConverterOptions
    {
        /// <summary>
        /// 匯入的 Table 是否含有表頭。
        /// </summary>
        public bool HasSchema { get; set; } = true;

        /// <summary>
        /// 決定欄位作為 Key 時的名稱。
        /// 輸入參數為 (欄位index, 欄位名稱)，輸出結果為 Key，
        /// 若 HasSchema 為 false 時，欄位名稱將是 null。
        /// 預設為 (欄位名稱 ?? 欄位index)。
        /// 若 return null，代表此欄位不匯入。
        /// </summary>
        public Func<int, string?, string?> CustomKeyNamePolicy { get; set; } = (index, fieldName) => fieldName ?? index.ToString();

        /// <summary>
        /// 決定將值寫入至 Vale 時應寫入的值。
        /// 輸入參數為 (Key, 原始值)，輸出結果為「應寫入的值」。
        /// 預設為原值。
        /// </summary>
        public Func<string, object?, object?> CustomValuePolicy { get; set; } = (key, value) => value;
    }
}
