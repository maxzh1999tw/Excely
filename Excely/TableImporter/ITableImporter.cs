namespace Excely.TableImporter
{
    /// <summary>
    /// 定義了一個可以將表格匯入為指定資料結構的執行單元。
    /// </summary>
    /// <typeparam name="TOutput">目標資料結構</typeparam>
    public interface ITableImporter<TOutput>
    {
        /// <summary>
        /// 將 Table 匯入為指定資料結構。
        /// </summary>
        /// <returns>匯入結果</returns>
        public TOutput Import(ExcelyTable table);
    }
}
