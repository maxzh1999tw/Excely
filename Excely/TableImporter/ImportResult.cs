namespace Excely.TableImporter
{
    public struct ImportResult<T>
    {
        public T ResultData;

        public Dictionary<CellLocation, Exception> CellErrors;

        public ImportResult(T resultData, Dictionary<CellLocation, Exception> cellErrors)
        {
            ResultData = resultData;
            CellErrors = cellErrors;
        }

        public readonly bool IsAllSuccess => !CellErrors.Any();
    }
}
