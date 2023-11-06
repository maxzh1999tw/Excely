using Excely.TableFactories;
using Excely.TableImporter;

namespace Excely.Workflows
{
    /// <summary>
    /// 提供從資料輸入到轉換為資料結構的完整工作流程。
    /// </summary>
    /// <typeparam name="TInput">資料的輸入型別</typeparam>
    public class ExcelyImporter<TInput, TOutput> : BaseWorkflow<TInput>
    {
        public ITableImporter<TOutput> TableImporter { get; set; }

        /// <summary>
        /// 以指定的 ITableFactory 物件初始化新的實例。
        /// </summary>
        public ExcelyImporter(ITableFactory<TInput> tableFactory, ITableImporter<TOutput> tableImporter) : base(tableFactory) 
        { 
            TableImporter = tableImporter;
        }

        /// <summary>
        /// 將指定的資料來源匯入為指定資料結構
        /// </summary>
        /// <param name="dataSource">資料來源</param>
        /// <returns>匯入結果</returns>
        public TOutput Import(TInput dataSource)
        {
            var table = TableFactory.GetTable(dataSource);
            return TableImporter.Import(table);
        }
    }
}

