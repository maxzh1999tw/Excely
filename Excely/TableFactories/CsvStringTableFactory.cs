using System.Text;

namespace Excely.TableFactories
{
    public class CsvStringTableFactory : ITableFactory<string>
    {
        public ExcelyTable GetTable(string sourceData)
        {
            var result = new List<IList<object?>>();
            var row = new List<object?>();
            StringBuilder colBuilder = new();

            bool isInQuotes = false;
            for (int i = 0; i < sourceData.Length; i++)
            {
                var currentChar = sourceData[i];

                if (isInQuotes)
                {
                    if (currentChar == '\"')
                    {
                        if (i < sourceData.Length - 1)
                        {
                            var nextChar = sourceData[i + 1];

                            // 拖逸
                            if (nextChar == '\"')
                            {
                                colBuilder.Append(currentChar);
                                i++;
                            }

                            // 換欄位
                            else if (nextChar == ',')
                            {
                                isInQuotes = false;
                                SaveCol(row, colBuilder);
                                i++;
                            }

                            // 換行
                            else if (nextChar == '\n')
                            {
                                isInQuotes = false;
                                SaveCol(row, colBuilder);
                                SaveRow(result, ref row);
                                i++;
                            }

                            // 換行
                            else if (i < sourceData.Length - 2 && sourceData.Substring(i + 1, 2) == "\r\n")
                            {
                                isInQuotes = false;
                                SaveCol(row, colBuilder);
                                SaveRow(result, ref row);
                                i += 2;
                            }

                            // 不合法
                            else
                            {
                                throw new InvalidCastException("欄位引用符號 '\\\n' 後面必須緊臨欄位分隔符號 ',' 或換行符。");
                            }
                        }
                        else break;
                    }
                    else
                    {
                        colBuilder.Append(currentChar);
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
                        SaveCol(row, colBuilder);
                    }

                    // 換行
                    else if (currentChar == '\n')
                    {
                        SaveCol(row, colBuilder);
                        SaveRow(result, ref row);
                    }

                    // 換行
                    else if (i < sourceData.Length - 2 && sourceData.Substring(i + 1, 2) == "\r\n")
                    {
                        SaveCol(row, colBuilder);
                        SaveRow(result, ref row);
                        i++;
                    }

                    else
                    {
                        colBuilder.Append(currentChar);
                    }
                }
            }

            if (colBuilder.Length > 0)
            {
                SaveCol(row, colBuilder);
            }
            if (row.Any())
            {
                SaveRow(result, ref row);
            }

            return new ExcelyTable(result);
        }

        private static void SaveCol(List<object?> row, StringBuilder colBuilder) {
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
