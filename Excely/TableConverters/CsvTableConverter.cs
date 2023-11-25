using System.Text;

namespace Excely.TableConverters
{
    /// <summary>
    /// 將 Table 轉換為 Csv 字串。
    /// </summary>
    /// <typeparam name="TOutput">轉出型別 (目前支援 string 與 MemoryStream)</typeparam>
    public class CsvStringTableConverter<TOutput> : ITableConverter<TOutput>
    {
        public TOutput ConvertFrom(ExcelyTable table)
        {
            var stringResult = CsvStringTableConverter<TOutput>.ConvertToString(table);

            // 轉換為 string
            if (typeof(TOutput) == typeof(string) || typeof(TOutput).IsAssignableFrom(typeof(string)))
            {
                return (TOutput)(object)stringResult;
            }

            // 轉換為 MemoryStream
            else if (typeof(TOutput) == typeof(MemoryStream) || typeof(TOutput).IsAssignableFrom(typeof(MemoryStream)))
            {
                var stream = new MemoryStream(Encoding.UTF8.GetBytes(stringResult));
                return (TOutput)(object)stream;
            }

            // 尚不支援其他轉換類型
            throw new NotSupportedException($"TOutput 為不支援的轉出型別。");
        }

        /// <summary>
        /// 將 Table 轉換為 Csv 字串的核心邏輯。
        /// </summary>
        /// <param name="table">來源 Table</param>
        /// <returns>Csv 字串</returns>
        private static string ConvertToString(ExcelyTable table)
        {
            StringBuilder stringBuilder = new();

            // 遍歷每 Row
            for (int rowIndex = 0; rowIndex < table.MaxRowCount; rowIndex++)
            {
                // 遍歷每 Column
                for (int columnIndex = 0; columnIndex < table.MaxColumnCount; columnIndex++)
                {
                    // 將資料加入 stringBuilder
                    stringBuilder.Append(table.Data[rowIndex][columnIndex]?.ToString());

                    // 若後面還有欄位，加上欄位分隔符
                    if (columnIndex != table.MaxColumnCount - 1)
                    {
                        stringBuilder.Append(',');
                    }
                }

                // 若後面還有資料，加上換行符
                if (rowIndex != table.MaxRowCount - 1)
                {
                    stringBuilder.Append(Environment.NewLine);
                }
            }

            return stringBuilder.ToString();
        }
    }
}
