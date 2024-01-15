using System.Formats.Asn1;
using System.Reflection;

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
        /// 決定 key 是否應作為欄位匯出。
        /// 預設為全部欄位都匯出。
        /// </summary>
        public KeyShowPolicyDelegate KeyShowPolicy { get; set; } = _ => true;

        /// <summary>
        /// 決定 key 作為欄位時的名稱。
        /// 預設為 key。
        /// </summary>
        public KeyNamePolicyDelegate KeyNamePolicy { get; set; } = key => key;

        /// <summary>
        /// 決定 key 作為欄位時的權重(越小越靠前)。
        /// 預設為 key 的預設順序。
        /// </summary>
        public KeyOrderPolicyDelegate KeyOrderPolicy { get; set; } = (allKeys, key) => Array.IndexOf(allKeys, key);

        /// <summary>
        /// 決定資料寫入欄位時的值。
        /// 預設為 Value。
        /// </summary>
        public CustomValuePolicyDelegate CustomValuePolicy { get; set; } = (key, dict) => dict.GetValueOrDefault(key, null);

        #region ===== Policy delegates =====

        /// <summary>
        /// 決定 key 是否應作為欄位匯出。
        /// </summary>
        /// <param name="key">當前決定的 key</param>
        /// <returns>是否應匯出</returns>
        public delegate bool KeyShowPolicyDelegate(string key);

        /// <summary>
        /// 決定 key 作為欄位時的名稱。
        /// </summary>
        /// <param name="key">當前決定的 key</param>
        /// <returns>欄位名稱</returns>
        public delegate string? KeyNamePolicyDelegate(string key);

        /// <summary>
        /// 決定 key 作為欄位時的權重(越小越靠前)。
        /// </summary>
        /// <param name="allKeys">所有 key</param>
        /// <param name="key">當前決定的 key</param>
        /// <returns>欄位權重</returns>
        public delegate int KeyOrderPolicyDelegate(string[] allKeys, string key);

        /// <summary>
        /// 決定資料寫入欄位時的值。
        /// </summary>
        /// <param name="key">當前決定的 key</param>
        /// <param name="writtingDict">當前正在寫入的 Dictionary</param>
        /// <returns>應寫入的值</returns>
        public delegate object? CustomValuePolicyDelegate(string key, Dictionary<string, object?> writtingDict);

        #endregion ===== Policy delegates =====
    }
}
