using Excely.EPPlus.LGPL.TableFactories;
using OfficeOpenXml;

namespace Excely.Workflows
{
	/// <summary>
	/// 提供快速建立從 Excel 到指定資料結構之 Importer 的方法。
	/// </summary>
	public class XlsxImporter : ExcelyImporterBase<ExcelWorksheet>
	{
		protected XlsxTableFactory XlsxTableFactory { get; set; } = new XlsxTableFactory();

		#region === 建構子 ==

		public XlsxImporter()
		{ }

		public XlsxImporter(CellLocation? startCell, CellLocation? endCell)
		{
			XlsxTableFactory = new XlsxTableFactory();
			if (startCell != null)
			{
				XlsxTableFactory.StartCell = startCell.Value;
			}
			XlsxTableFactory.EndCell = endCell;
		}

		#endregion === 建構子 ==

		protected override ExcelyTable GetTable(ExcelWorksheet input) => XlsxTableFactory.GetTable(input);

		protected override ExcelWorksheet GetDataSource(string filePath)
		{
			var package = new ExcelPackage(new FileInfo(filePath));
			return package.Workbook.Worksheets[0];
		}
	}
}