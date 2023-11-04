namespace Excely.TableImporter
{
    public struct ImportResult<T>
    {
        public T ResultData;

        public Dictionary<CellLocation, Exception> ErrorCells;

        public ImportResult(T resultData, Dictionary<CellLocation, Exception> errorCells)
        {
            ResultData = resultData;
            ErrorCells = errorCells;
        }

        public readonly bool IsAllSuccess => !ErrorCells.Any();
    }
}
