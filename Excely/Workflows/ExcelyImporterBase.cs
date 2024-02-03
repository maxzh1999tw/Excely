using Excely.TableConverters;

namespace Excely.Workflows
{
    /// <summary>
    /// 提供從資料輸入到轉換為資料結構的完整工作流程之基底類別。
    /// </summary>
    /// <typeparam name="TInput">資料的輸入型別</typeparam>
    public abstract class ExcelyImporterBase<TInput>
    {
        protected abstract ExcelyTable GetTable(TInput input);

        public List<Action<ExcelyTable>> DoAfterGetTableCallbackList { get; set; } = new List<Action<ExcelyTable>>();

        /// <summary>
        /// 將資料匯入為物件列表。
        /// </summary>
        /// <typeparam name="TClass">目標類別</typeparam>
        /// <param name="dataSource">資料來源</param>
        /// <param name="options">匯入邏輯</param>
        /// <returns>匯入結果</returns>
        public IEnumerable<TClass> ToClassList<TClass>(
            TInput dataSource,
            ClassListTableConverterOptions<TClass>? options = null)
            where TClass : class, new()
        {
            var table = GetTable(dataSource);
            DoAfterGetTableCallbackList.ForEach(x => x.Invoke(table));
            var converter = options == null ? new ClassListTableConverter<TClass>() : new ClassListTableConverter<TClass>(options);
            return converter.ConvertFrom(table);
        }

        /// <summary>
        /// 將資料匯入為字典列表。
        /// </summary>=
        /// <param name="dataSource">資料來源</param>
        /// <param name="options">匯入邏輯</param>
        /// <returns>匯入結果</returns>
        public IEnumerable<Dictionary<string, object?>> ToDictionaryList(
            TInput dataSource,
            DictionaryListTableConverterOptions? options = null)
        {
            var table = GetTable(dataSource);
            DoAfterGetTableCallbackList.ForEach(x => x.Invoke(table));
            var converter = options == null ? new DictionaryListTableConverter() : new DictionaryListTableConverter(options);
            return converter.ConvertFrom(table);
        }
    }
}

