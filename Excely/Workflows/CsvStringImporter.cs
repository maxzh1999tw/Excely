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

		public CsvStringImporter()
		{ }

		#endregion === 建構子 ==

		protected override ExcelyTable GetTable(string input) => CsvTableFactory.GetTable(input);

		protected override string GetDataSource(string filePath)
		{
			using var streamReader = new StreamReader(filePath);
			return streamReader.ReadToEnd();
		}
	}
}