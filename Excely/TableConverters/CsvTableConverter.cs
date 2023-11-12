using System.Text;

namespace Excely.TableConverters
{
    public class CsvStringTableConverter<TOutput> : ITableConverter<TOutput>
    {
        public TOutput Convert(ExcelyTable table)
        {
            var stringResult = CsvStringTableConverter<TOutput>.ConvertToString(table);
            if (typeof(TOutput) == typeof(string) || typeof(TOutput).IsAssignableFrom(typeof(string)))
            {
                return (TOutput)(object)stringResult;
            }
            else if (typeof(TOutput) == typeof(MemoryStream) || typeof(TOutput).IsAssignableFrom(typeof(MemoryStream)))
            {
                var stream = new MemoryStream(Encoding.UTF8.GetBytes(stringResult));
                return (TOutput)(object)stream;
            }

            throw new NotSupportedException();
        }

        private static string ConvertToString(ExcelyTable table)
        {
            StringBuilder stringBuilder = new();
            for (int rowIndex = 0; rowIndex < table.MaxRowCount; rowIndex++)
            {
                for (int colIndex = 0; colIndex < table.MaxColCount; colIndex++)
                {
                    stringBuilder.Append(table.Data[rowIndex][colIndex]?.ToString());
                    if (colIndex != table.MaxColCount - 1)
                    {
                        stringBuilder.Append(',');
                    }
                }
                if (rowIndex != table.MaxRowCount - 1)
                {
                    stringBuilder.Append(Environment.NewLine);
                }
            }

            return stringBuilder.ToString();
        }
    }
}
