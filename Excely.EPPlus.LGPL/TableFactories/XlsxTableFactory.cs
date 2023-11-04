using Excely.TableFactories;
using OfficeOpenXml;

namespace Excely.EPPlus.LGPL.TableFactories
{
    public class XlsxTableFactory : ITableFactory<ExcelWorksheet>
    {
        public CellLocation StartCell { get; set; } = new CellLocation(0, 0);

        public CellLocation? EndCell { get; set; }

        public ExcelyTable GetTable(ExcelWorksheet input)
        {
            CellLocation realEndCell;
            if (EndCell == null)
            {
                realEndCell = new CellLocation(input.Dimension.End.Row, input.Dimension.End.Column);
            }
            else
            {
                realEndCell = new CellLocation(
                    Math.Min(EndCell.Value.Row, input.Dimension.End.Row),
                    Math.Min(EndCell.Value.Column, input.Dimension.End.Column));
            }

            List<IList<object?>> result = new();
            for (int r = StartCell.Row + 1; r <= realEndCell.Row; r++)
            {
                List<object?> row = new();
                for (int c = StartCell.Column + 1; c <= realEndCell.Column; c++)
                {
                    row.Add(input.Cells[r, c].Value);
                }
                result.Add(row);
            }

            return new ExcelyTable(result);
        }
    }
}
