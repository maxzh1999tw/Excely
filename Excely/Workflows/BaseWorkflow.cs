using Excely.TableFactories;

namespace Excely.Workflows
{
    /// <summary>
    /// 提供匯入與匯出的基礎工作流程
    /// </summary>
    /// <typeparam name="TInput">資料的輸入型別</typeparam>
    public class BaseWorkflow<TInput>
    {
        /// <summary>
        /// 將輸入資料轉換為 Table 的 ITableFactory 物件。
        /// </summary>
        protected ITableFactory<TInput> TableFactory { get; set; }

        /// <summary>
        /// 以指定的 ITableFactory 物件初始化新的實例。
        /// </summary>
        public BaseWorkflow(ITableFactory<TInput> tableFactory)
        {
            TableFactory = tableFactory;
        }

        /// <summary>
        /// 將輸入資料轉換為 Table。
        /// </summary>
        /// <param name="sourceData">輸入資料</param>
        /// <returns>ExcelyTable 格式的資料</returns>
        public ExcelyTable GetTable(TInput sourceData) => TableFactory.GetTable(sourceData);
    }
}
