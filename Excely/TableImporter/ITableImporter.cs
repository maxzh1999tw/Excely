namespace Excely.TableImporter
{
    /// <summary>
    /// 定義了一個可以將表格匯入為指定資料結構的執行單元。
    /// </summary>
    /// <typeparam name="TInput">目標資料結構</typeparam>
    public interface ITableImporter<TInput>
    {
        /// <summary>
        /// 將 Table 匯入為指定資料結構。
        /// </summary>
        /// <returns>匯入結果</returns>
        public ImportResult<TInput> Import(ExcelyTable table);
    }
}
