namespace Excely.TableFactories
{
    public class CsvStringTableFactory : ITableFactory<string>
    {
        public ExcelyTable GetTable(string sourceData)
        {
            var result = new List<IList<object?>>();
            var rows = sourceData.Split(Environment.NewLine);
            foreach (var row in rows)
            {
                var rowData = new List<object?>();
                var cols = row.Split(',');
                foreach (var col in cols)
                {
                    rowData.Add(col);
                }
                result.Add(rowData);
            }
            return new ExcelyTable(result);
        }
    }
}
