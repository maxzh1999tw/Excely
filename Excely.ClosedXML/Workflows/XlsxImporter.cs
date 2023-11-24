using ClosedXML.Excel;
using Excely.ClosedXML.TableFactories;
using Excely.Workflows;

namespace Excely.ClosedXML.Workflows
{
    /// <summary>
    /// 提供快速建立從 Excel 到指定資料結構之 Importer 的方法。
    /// </summary>
    public class XlsxImporter : ExcelyImporter<IXLWorksheet>
    {
        protected XlsxTableFactory XlsxTableFactory { get; set; } = new XlsxTableFactory();

        #region === 建構子 ==
        public XlsxImporter() { }

        public XlsxImporter(CellLocation? startCell, CellLocation? endCell)
        {
            XlsxTableFactory = new XlsxTableFactory();
            if (startCell != null)
            {
                XlsxTableFactory.StartCell = startCell.Value;
            }
            XlsxTableFactory.EndCell = endCell;
        }
        #endregion

        protected override ExcelyTable GetTable(IXLWorksheet input) => XlsxTableFactory.GetTable(input);
    }
}
