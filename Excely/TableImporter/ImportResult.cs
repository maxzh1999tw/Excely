namespace Excely.TableImporter
{
    /// <summary>
    /// 代表匯入結果的結構。
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public struct ImportResult<T>
    {
        /// <summary>
        /// 匯入結果。
        /// </summary>
        public T ResultData;

        /// <summary>
        /// 匯入時出錯的 Cell 位置，與其錯誤。
        /// </summary>
        public Dictionary<CellLocation, Exception> CellErrors;

        public ImportResult(T resultData, Dictionary<CellLocation, Exception> cellErrors)
        {
            ResultData = resultData;
            CellErrors = cellErrors;
        }

        /// <summary>
        /// 是否全部匯入成功。
        /// </summary>
        public readonly bool IsAllSuccess => !CellErrors.Any();
    }
}
