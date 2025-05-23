using Excely.Shaders;
using Excely.TableConverters;
using Excely.TableFactories;

namespace Excely.Workflows
{
    /// <summary>
    /// 提供從資料輸入到匯出的完整工作流程。
    /// </summary>
    /// <typeparam name="TInput">資料的輸入型別</typeparam>
    public class ExcelyExporter<TInput>
    {
        /// <summary>
        /// 將輸入資料轉換為 Table 的 ITableFactory 物件。
        /// </summary>
        protected ITableFactory<TInput> TableFactory { get; set; }

        /// <summary>
        /// 取得匯出結果後依序執行的 IShaders。
        /// </summary>
        public IEnumerable<IShader> Shaders { get; set; } = Enumerable.Empty<IShader>();

        /// <summary>
        /// 以指定的 ITableFactory 物件初始化新的實例。
        /// </summary>
        public ExcelyExporter(ITableFactory<TInput> tableFactory)
        {
            TableFactory = tableFactory;
        }

        /// <summary>
        /// 將輸入資料轉換為 Table。
        /// </summary>
        /// <param name="sourceData">輸入資料</param>
        /// <returns>ExcelyTable 格式的資料</returns>
        public ExcelyTable GetTable(TInput sourceData) => TableFactory.GetTable(sourceData);

        /// <summary>
        /// 將輸入資料轉換為 Csv 字串。
        /// </summary>
        /// <param name="sourceData">輸入資料</param>
        /// <returns>Csv 字串</returns>
        public string ToCsvString(TInput sourceData)
        {
            var table = GetTable(sourceData);
            var tableWriter = new CsvStringTableConverter<string>();
            var result = tableWriter.ConvertFrom(table);
            foreach (var shaders in Shaders)
            {
                result = shaders.Execute(result);
            }
            return result;
        }

        /// <summary>
        /// 將輸入資料轉換為檔案。
        /// </summary>
        /// <param name="sourceData">輸入資料</param>
        /// <returns>匯出之檔案</returns>
        public MemoryStream ToCsvFile(TInput sourceData)
        {
            var table = GetTable(sourceData);
            var tableWriter = new CsvStringTableConverter<MemoryStream>();
            var result = tableWriter.ConvertFrom(table);
            foreach (var shaders in Shaders)
            {
                result = shaders.Execute(result);
            }
            return result;
        }
    }

    /// <summary>
    /// 提供用於快速建立 ExcelyExporter 的靜態方法。
    /// </summary>
    public static class ExcelyExporter
    {
        /// <summary>
        /// 建立一個以 Class list 為輸入資料型別的 ExcelyExporter。
        /// </summary>
        /// <typeparam name="TClass">欲轉換的 Class</typeparam>
        public static ExcelyExporter<IEnumerable<TClass>> FromClassList<TClass>(
            ClassListTableFactoryOptions<TClass>? options = null,
            IEnumerable<IShader>? shaders = null)
            where TClass : class
        {
            var tableFactory = options == null ? new ClassListTableFactory<TClass>() : new ClassListTableFactory<TClass>(options);

            var exporter = new ExcelyExporter<IEnumerable<TClass>>(tableFactory);
            if (shaders != null)
            {
                exporter.Shaders = shaders;
            }
            return exporter;
        }

        /// <summary>
        /// 建立一個以 Class list 為輸入資料型別的 ExcelyExporter。
        /// </summary>
        /// <typeparam name="TClass">欲轉換的 Class</typeparam>
        public static ExcelyExporter<IEnumerable<Dictionary<string, object?>>> FromDictionaryList(
            DictionaryListTableFactoryOptions? options = null,
            IEnumerable<IShader>? shaders = null)
        {
            var tableFactory = options == null ? new DictionaryListTableFactory() : new DictionaryListTableFactory(options);

            var exporter = new ExcelyExporter<IEnumerable<Dictionary<string, object?>>>(tableFactory);
            if (shaders != null)
            {
                exporter.Shaders = shaders;
            }
            return exporter;
        }
    }
}
