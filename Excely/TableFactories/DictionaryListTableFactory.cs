namespace Excely.TableFactories
{
    /// <summary>
    /// 提供以字典 Key 為欄位，將字典集合傾印至表格的功能。
    /// </summary>
    public class DictionaryListTableFactory : ITableFactory<IEnumerable<Dictionary<string, object?>>>
    {
        /// <summary>
        /// 轉換過程的執行細節。
        /// </summary>
        protected DictionaryListTableFactoryOptions Options { get; set; } = new DictionaryListTableFactoryOptions();

        #region === 建構子 ===
        public DictionaryListTableFactory() { }

        public DictionaryListTableFactory(DictionaryListTableFactoryOptions options)
        {
            Options = options;
        }
        #endregion

        public ExcelyTable GetTable(IEnumerable<Dictionary<string, object?>> sourceData)
        {
            var keys = sourceData
                .SelectMany(x => x.Keys)
                .Distinct()
                .Where(x => Options.KeyShowPolicy(x))
                .ToArray();
            keys = keys.OrderBy(x => Options.KeyOrderPolicy(keys, x)).ToArray();

            var table = new List<IList<object?>>();

            if (Options.WithSchema)
            {
                table.Add(GetSchema(keys));
            }

            foreach (var item in sourceData)
            {
                table.Add(GetRow(item, keys));
            }

            return new ExcelyTable(table);
        }

        /// <summary>
        /// 將 Key 轉換為表頭。
        /// </summary>
        /// <param name="keys">欲匯出的 Keys</param>
        /// <returns>表頭列</returns>
        private IList<object?> GetSchema(string[] keys)
        {
            return keys.Select(k => Options.KeyNamePolicy(k)).ToList<object?>();
        }

        /// <summary>
        /// 將字典轉換為資料列。
        /// </summary>
        /// <param name="item">來源字典</param>
        /// <param name="keys">欲匯出的 Keys</param>
        /// <returns>資料列</returns>
        private IList<object?> GetRow(Dictionary<string, object?> item, string[] keys)
        {
            return keys.Select(k => Options.CustomValuePolicy(k, item)).ToList();
        }
    }

    public class DictionaryListTableFactoryOptions
    {
        /// <summary>
        /// 決定匯出時是否帶有表頭。
        /// 預設為是。
        /// </summary>
        public bool WithSchema { get; set; } = true;

        /// <summary>
        /// 決定 key 是否應作為欄位匯出的執行邏輯。
        /// 輸入參數為 key，輸出結果為「是否應作為欄位匯出」，
        /// 預設為全部欄位都匯出。
        /// </summary>
        public Func<string, bool> KeyShowPolicy { get; set; } = _ => true;

        /// <summary>
        /// 決定 key 作為欄位時的名稱。
        /// 輸入參數為 key，輸出結果為「欄位名稱」，
        /// 預設為 key。
        /// </summary>
        public Func<string, string?> KeyNamePolicy { get; set; } = k => k;

        /// <summary>
        /// 決定 key 作為欄位時的順序。
        /// 輸入參數為 (所有key(預設排序), key)，輸出結果為「排序(由小到大)」，
        /// 預設為依 key 出現順序排序。
        /// </summary>
        public Func<string[], string, int> KeyOrderPolicy { get; set; } = (keys, k) => Array.IndexOf(keys, k);

        /// <summary>
        /// 決定資料寫入欄位時的值。
        /// 輸入參數為 (key, 當前匯出物件)，輸出結果為「欲寫入欄位的值」，
        /// 預設為 Value。
        /// </summary>
        public Func<string, Dictionary<string, object?>, object?> CustomValuePolicy { get; set; } = (k, dict) => dict.GetValueOrDefault(k, null);
    }
}
