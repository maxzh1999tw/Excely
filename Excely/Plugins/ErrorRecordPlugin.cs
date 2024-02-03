namespace Excely.Plugins
{
    public class ErrorRecordPlugin
    {
        public ExcelyTable? Table { get; set; }

        public Dictionary<CellLocation, string> Errors { get; set; } = new Dictionary<CellLocation, string>();

        public void DoAfterGetTable(ExcelyTable table)
        {
            Table = table;
        }

        public void RecordError(CellLocation location, string errorMessage)
        {
            Errors[location] = errorMessage;
        }
    }
}
