namespace Excely.TableImporter
{
    public interface ITableImporter<T>
    {
        public ImportResult<T> Import(ExcelyTable table);
    }
}
