using Excely.TableFactories;

namespace Excely.Workflows
{
    /// <summary>
    /// 以 Csv 字串為來源資料的 Importer。
    /// </summary>
    public class CsvStringImporter : ExcelyImporterBase<string>
    {
        protected CsvStringTableFactory CsvTableFactory { get; set; } = new();

        #region === 建構子 ==
        public CsvStringImporter() { }
        #endregion

        protected override ExcelyTable GetTable(string input) => CsvTableFactory.GetTable(input);
    }
}
