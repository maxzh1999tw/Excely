namespace Excely.TableFactorys
{
    /// <summary>
    /// 定義了一個可以取得表格的執行單元
    /// </summary>
    /// <typeparam name="TInput"></typeparam>
    public interface ITableFactory<TInput>
    {
        /// <summary>
        /// 將指定的物件集合傾印至表格
        /// </summary>
        /// <param name="sourceData">來源資料</param>
        /// <returns>表格</returns>
        ExcelyTable GetTable(TInput sourceData);
    }
}
