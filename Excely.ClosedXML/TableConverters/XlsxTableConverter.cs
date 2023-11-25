using ClosedXML.Excel;
using Excely.TableConverters;

namespace Excely.ClosedXML.TableConverters
{
    /// <summary>
    /// 提供將表格匯出為 Xlsx 的功能。
    /// </summary>
    public class XlsxTableConverter : ITableConverter<IXLWorksheet>
    {
        /// <summary>
        /// 起始匯出儲存格。
        /// </summary>
        public CellLocation StartCell { get; set; } = new(0, 0);

        /// <summary>
        /// 目標工作表。
        /// </summary>
        public IXLWorksheet TargetWorksheet { get; set; }

        public XlsxTableConverter(IXLWorksheet targetWorksheet)
        {
            TargetWorksheet = targetWorksheet;
        }

        public XlsxTableConverter(IXLWorksheet targetWorksheet, CellLocation startCell)
        {
            TargetWorksheet = targetWorksheet;
            StartCell = startCell;
        }

        public IXLWorksheet ConvertFrom(ExcelyTable table)
        {
            int rowPointer = StartCell.Row;
            foreach (var rowData in table.Data)
            {
                int colPointer = StartCell.Column;
                foreach (var colData in rowData)
                {
                    TargetWorksheet.Cell(rowPointer + 1, colPointer + 1).Value = XLCellValue.FromObject(colData);
                    colPointer++;
                }
                rowPointer++;
            }

            return TargetWorksheet;
        }
    }
}
