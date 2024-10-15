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

		public DictionaryListTableFactory()
		{ }

		public DictionaryListTableFactory(DictionaryListTableFactoryOptions options)
		{
			Options = options;
		}

		#endregion === 建構子 ===

		public ExcelyTable GetTable(IEnumerable<Dictionary<string, object?>> sourceData)
		{
			var keys = sourceData
				.SelectMany(x => x.Keys)
				.Distinct()
				.Where(x => Options.KeyShowPolicy(new() { Key = x }))
				.ToArray();
			keys = keys.OrderBy(x => Options.KeyOrderPolicy(new()
			{
				AllKeys = keys,
				Key = x
			})).ToArray();

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
			return keys.Select(k => Options.KeyNamePolicy(new() { Key = k })).ToList<object?>();
		}

		/// <summary>
		/// 將字典轉換為資料列。
		/// </summary>
		/// <param name="item">來源字典</param>
		/// <param name="keys">欲匯出的 Keys</param>
		/// <returns>資料列</returns>
		private IList<object?> GetRow(Dictionary<string, object?> item, string[] keys)
		{
			return keys.Select(k => Options.CustomValuePolicy(new()
			{
				Key = k,
				WrittingDict = item
			})).ToList();
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
		public KeyShowPolicyDelegate KeyShowPolicy { get; set; }

		/// <summary>
		/// 決定 key 作為欄位時的名稱。
		/// 預設為 key。
		/// </summary>
		public KeyNamePolicyDelegate KeyNamePolicy { get; set; }

		/// <summary>
		/// 決定 key 作為欄位時的權重(越小越靠前)。
		/// 預設為 key 的預設順序。
		/// </summary>
		public KeyOrderPolicyDelegate KeyOrderPolicy { get; set; }

		/// <summary>
		/// 決定資料寫入欄位時的值。
		/// 預設為 Value。
		/// </summary>
		public CustomValuePolicyDelegate CustomValuePolicy { get; set; }

		public DictionaryListTableFactoryOptions()
		{
			KeyShowPolicy = DefaultKeyShowPolicy;
			KeyNamePolicy = DefaultKeyNamePolicy;
			KeyOrderPolicy = DefaultKeyOrderPolicy;
			CustomValuePolicy = DefaultCustomValuePolicy;
		}

		#region ===== Policy delegates =====

		/// <summary>
		/// 決定 key 是否應作為欄位匯出。
		/// </summary>
		/// <returns>是否應匯出</returns>
		public delegate bool KeyShowPolicyDelegate(KeyShowPolicyDelegateParam param);

		public struct KeyShowPolicyDelegateParam
		{
			/// <summary>
			/// 當前決定的 key
			/// </summary>
			public string Key;
		}

		/// <summary>
		/// 決定 key 作為欄位時的名稱。
		/// </summary>
		/// <returns>欄位名稱</returns>
		public delegate string? KeyNamePolicyDelegate(KeyNamePolicyDelegateParam param);

		public struct KeyNamePolicyDelegateParam
		{
			/// <summary>
			/// 當前決定的 key
			/// </summary>
			public string Key;
		}

		/// <summary>
		/// 決定 key 作為欄位時的權重(越小越靠前)。
		/// </summary>
		/// <returns>欄位權重</returns>
		public delegate int KeyOrderPolicyDelegate(KeyOrderPolicyDelegateParam param);

		public struct KeyOrderPolicyDelegateParam
		{
			/// <summary>
			/// 所有 key
			/// </summary>
			public string[] AllKeys;

			/// <summary>
			/// 當前決定的 key
			/// </summary>
			public string Key;
		}

		/// <summary>
		/// 決定資料寫入欄位時的值。
		/// </summary>
		/// <returns>應寫入的值</returns>
		public delegate object? CustomValuePolicyDelegate(CustomValuePolicyDelegateParam param);

		public struct CustomValuePolicyDelegateParam
		{
			/// <summary>
			/// 當前決定的 key
			/// </summary>
			public string Key;

			/// <summary>
			/// 當前正在寫入的 Dictionary
			/// </summary>
			public Dictionary<string, object?> WrittingDict;
		}

		#endregion ===== Policy delegates =====

		#region ===== Default policies =====

		public static bool DefaultKeyShowPolicy(KeyShowPolicyDelegateParam _) => true;

		public static string? DefaultKeyNamePolicy(KeyNamePolicyDelegateParam param) => param.Key;

		public static int DefaultKeyOrderPolicy(KeyOrderPolicyDelegateParam param) => Array.IndexOf(param.AllKeys, param.Key);

		public static object? DefaultCustomValuePolicy(CustomValuePolicyDelegateParam param) => param.WrittingDict.GetValueOrDefault(param.Key, null);

		#endregion ===== Default policies =====
	}
}