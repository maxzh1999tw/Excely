using System.Text;

namespace Excely.TableFactories
{
    /// <summary>
    /// 將 Csv 字串轉換為 Table。
    /// </summary>
    public class CsvStringTableFactory : ITableFactory<string>
    {
        public ExcelyTable GetTable(string sourceData)
        {
            var result = new List<IList<object?>>();
            var row = new List<object?>();
            StringBuilder columnBuilder = new();

            bool isInQuotes = false;

            // 遍歷字元
            for (int i = 0; i < sourceData.Length; i++)
            {
                var currentChar = sourceData[i];

                // 若在引號欄位中
                if (isInQuotes)
                {
                    // 遇到 "，不是跳脫字元 "" 就是欄位結尾
                    if (currentChar == '\"')
                    {
                        if (i < sourceData.Length - 1)
                        {
                            var nextChar = sourceData[i + 1];

                            // 拖逸
                            if (nextChar == '\"')
                            {
                                columnBuilder.Append(currentChar);
                                i++;
                            }

                            // 換欄位
                            else if (nextChar == ',')
                            {
                                isInQuotes = false;
                                SaveColumn(row, columnBuilder);
                                i++;
                            }

                            // 換行
                            else if (nextChar == '\n')
                            {
                                isInQuotes = false;
                                SaveColumn(row, columnBuilder);
                                SaveRow(result, ref row);
                                i++;
                            }

                            // 換行
                            else if (i < sourceData.Length - 2 && sourceData.Substring(i + 1, 2) == "\r\n")
                            {
                                isInQuotes = false;
                                SaveColumn(row, columnBuilder);
                                SaveRow(result, ref row);
                                i += 2;
                            }

                            // 不合法
                            else
                            {
                                throw new InvalidCastException("欄位引用符號 (\") 必須代表欄位結尾，或跳脫字元組 (\"\")。");
                            }
                        }
                        // 此為欄位結尾，同時也是字串結尾
                        else
                        {
                            isInQuotes = false;
                        }
                    }

                    // 其餘字元一律視為文字
                    else
                    {
                        columnBuilder.Append(currentChar);
                    }
                }
                else
                {
                    // 進 Quotes 欄位
                    if (currentChar == '\"')
                    {
                        isInQuotes = true;
                    }

                    // 換欄位
                    else if (currentChar == ',')
                    {
                        SaveColumn(row, columnBuilder);
                    }

                    // 換行
                    else if (currentChar == '\n')
                    {
                        SaveColumn(row, columnBuilder);
                        SaveRow(result, ref row);
                    }

                    // 換行
                    else if (i < sourceData.Length - 1 && sourceData.Substring(i, 2) == "\r\n")
                    {
                        SaveColumn(row, columnBuilder);
                        SaveRow(result, ref row);
                        i++;
                    }

                    // 其餘字元一律視為文字
                    else
                    {
                        columnBuilder.Append(currentChar);
                    }
                }
            }

            // 字串結束後必須不在引號中
            if (isInQuotes)
            {
                throw new InvalidCastException("Csv 格式錯誤：未閉合的引號。");
            }

            // 最後一欄與最後一列可能沒有被加入到結果中，補上
            if (columnBuilder.Length > 0)
            {
                SaveColumn(row, columnBuilder);
            }
            if (row.Any())
            {
                SaveRow(result, ref row);
            }

            // 刪除空行
            if (result.Any(row => row.Count > 1))
            {
                result = result.Where(x => x.Count != 1).ToList();
            }

            return new ExcelyTable(result.Where(x => x.Any()).ToArray());
        }

        private static void SaveColumn(List<object?> row, StringBuilder colBuilder)
        {
            row.Add(colBuilder.ToString());
            colBuilder.Clear();
        }

        private static void SaveRow(List<IList<object?>> result, ref List<object?> row)
        {
            result.Add(row);
            row = new List<object?>();
        }
    }
}
