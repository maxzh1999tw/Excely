using Excely.TableFactories;

namespace Excely.Workflows
{
    public class CsvStringImporter : ExcelyImporter<string>
    {
        protected CsvStringTableFactory CsvTableFactory { get; set; } = new();

        #region === 建構子 ==
        public CsvStringImporter() { }
        #endregion

        protected override ExcelyTable GetTable(string input) => CsvTableFactory.GetTable(input);
    }
}
