namespace Excely
{
    /// <summary>
    /// 表格資料結構。
    /// </summary>
    public class ExcelyTable
    {
        /// <summary>
        /// 表格資料。
        /// </summary>
        public IList<IList<object?>> Data { get; init; }

        /// <summary>
        /// 表格高度(含表頭)。
        /// </summary>
        public int MaxRowCount => Data.Count;

        /// <summary>
        /// 表格寬度。
        /// </summary>
        public int MaxColumnCount => Data.OrderBy(x => x.Count).FirstOrDefault()?.Count ?? 0;

        public ExcelyTable(IList<IList<object?>> data)
        {
            Data = data;
        }
    }
}
