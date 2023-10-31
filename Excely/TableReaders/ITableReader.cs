namespace Excely.TableReaders
{
    public interface ITableReader<TInput>
    {
        public ExcelyTable Read(TInput input);
    }
}
