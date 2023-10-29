namespace Excely.TableWriters
{
    /// <summary>
    /// 定義了可以將表格匯出至特定目標的執行單元
    /// </summary>
    /// <typeparam name="TOutput">匯出目標型別</typeparam>
    public interface ITableWriter<TOutput>
    {
        /// <summary>
        /// 將表格以指定的目標型別匯出
        /// </summary>
        /// <param name="table">來源資料</param>
        /// <returns>匯出結果</returns>
        TOutput Write(ExcelyTable table);
    }
}
