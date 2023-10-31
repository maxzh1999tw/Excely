using Excely.TableWriters;
using OfficeOpenXml;

namespace Excely.EPPlus.LGPL.TableWriters
{
    /// <summary>
    /// 提供將表格匯出為 Xlsx 的功能
    /// </summary>
    public class XlsxTableWriter : ITableWriter<ExcelWorksheet>
    {
        /// <summary>
        /// 起始匯出儲存格
        /// </summary>
        public (int Row, int Col) StartCell { get; set; } = (0, 0);

        /// <summary>
        /// 目標工作表
        /// </summary>
        public ExcelWorksheet TargetWorksheet { get; set; }

        public XlsxTableWriter(ExcelWorksheet targetWorksheet)
        {
            TargetWorksheet = targetWorksheet;
        }

        public XlsxTableWriter(ExcelWorksheet targetWorksheet, (int Row, int Col) startCell)
        {
            TargetWorksheet = targetWorksheet;
            StartCell = startCell;
        }

        public ExcelWorksheet Write(ExcelyTable table)
        {
            int rowPointer = StartCell.Row;
            foreach (var rowData in table.Data)
            {
                int colPointer = StartCell.Col;
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
