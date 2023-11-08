using Excely.TableConverters;
using OfficeOpenXml;

namespace Excely.EPPlus.LGPL.TableConverters
{
    /// <summary>
    /// 提供將表格匯出為 Xlsx 的功能。
    /// </summary>
    public class XlsxTableConverter : ITableConverter<ExcelWorksheet>
    {
        /// <summary>
        /// 起始匯出儲存格。
        /// </summary>
        public CellLocation StartCell { get; set; } = new(0, 0);

        /// <summary>
        /// 目標工作表。
        /// </summary>
        public ExcelWorksheet TargetWorksheet { get; set; }

        public XlsxTableConverter(ExcelWorksheet targetWorksheet)
        {
            TargetWorksheet = targetWorksheet;
        }

        public XlsxTableConverter(ExcelWorksheet targetWorksheet, CellLocation startCell)
        {
            TargetWorksheet = targetWorksheet;
            StartCell = startCell;
        }

        public ExcelWorksheet Convert(ExcelyTable table)
        {
            int rowPointer = StartCell.Row;
            foreach (var rowData in table.Data)
            {
                int colPointer = StartCell.Column;
                foreach (var colData in rowData)
                {
                    TargetWorksheet.Cells[rowPointer + 1, colPointer + 1].Value = colData;
                    colPointer++;
                }
                rowPointer++;
            }

            return TargetWorksheet;
        }
    }
}
