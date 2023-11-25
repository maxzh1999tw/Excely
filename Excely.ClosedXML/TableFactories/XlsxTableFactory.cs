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
                realEndCell = new CellLocation(input.LastRowUsed().RowNumber(), input.LastColumnUsed().ColumnNumber());
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
                    var cell = input.Cell(r, c);
                    var cellValue = cell.Value;
                    object? value;

                    if (cellValue.Type == XLDataType.Blank)
                    {
                        value = null;
                    }
                    else if (cellValue.Type == XLDataType.Text)
                    {
                        value = cellValue.GetText();
                    }
                    else if (cellValue.Type == XLDataType.Number)
                    {
                        if (cell.TryGetValue<decimal>(out var number))
                        {
                            value = number;
                        }
                        else
                        {
                            throw new InvalidCastException();
                        }
                    }
                    else if (cellValue.Type == XLDataType.Boolean)
                    {
                        value = cellValue.GetBoolean();
                    }
                    else if (cellValue.Type == XLDataType.DateTime)
                    {
                        value = cellValue.GetDateTime();
                    }
                    else if (cellValue.Type == XLDataType.TimeSpan)
                    {
                        value = cellValue.GetTimeSpan();
                    }
                    else
                    {
                        throw new NotImplementedException();
                    }

                    row.Add(value);
                }
                result.Add(row);
            }

            return new ExcelyTable(result);
        }
    }
}
