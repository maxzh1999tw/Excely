using ClosedXML.Excel;
using Excely.TableFactories;

namespace Excely.ClosedXML.TableFactories
{
    /// <summary>
    /// 提供將 Xlsx 匯入為 Table 的功能。
    /// </summary>
    public class XlsxTableFactory : ITableFactory<IXLWorksheet>
    {
        /// <summary>
        /// 匯入起始儲存格。
        /// </summary>
        public CellLocation StartCell { get; set; } = new CellLocation(0, 0);

        /// <summary>
        /// 匯入結束儲存格。
        /// 若為 null 代表不設限。
        /// </summary>
        public CellLocation? EndCell { get; set; }

        public ExcelyTable GetTable(IXLWorksheet input)
        {
            CellLocation realEndCell;
            if (EndCell == null)
            {
                realEndCell = new CellLocation(input.RowCount(), input.ColumnCount());
            }
            else
            {
                realEndCell = new CellLocation(
                    Math.Min(EndCell.Value.Row, input.RowCount()),
                    Math.Min(EndCell.Value.Column, input.ColumnCount()));
            }

            List<IList<object?>> result = new();
            for (int r = StartCell.Row + 1; r <= realEndCell.Row; r++)
            {
                List<object?> row = new();
                for (int c = StartCell.Column + 1; c <= realEndCell.Column; c++)
                {
                    row.Add(input.Cell(r, c).Value);
                }
                result.Add(row);
            }

            return new ExcelyTable(result);
        }
    }
}
