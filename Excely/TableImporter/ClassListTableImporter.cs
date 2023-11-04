using System.Reflection;

namespace Excely.TableImporter
{
    public class ClassListTableImporter<T> : ITableImporter<IEnumerable<T>> where T : class, new()
    {
        public bool HasSchema { get; set; } = true;

        public bool StopWhenError { get; set; } = true;

        public Func<PropertyInfo, string?>? PropertyNamePolicy { get; set; }

        public Func<PropertyInfo, int?>? PropertyIndexPolicy { get; set; }

        public Func<PropertyInfo, object?, object?>? CustomValuePolicy { get; set; }

        private PropertyInfo[]? _TProperies;

        private PropertyInfo[] TProperties
        {
            get
            {
                _TProperies ??= typeof(T).GetProperties();
                return _TProperies;
            }
        }

        public ImportResult<IEnumerable<T>> Import(ExcelyTable table)
        {
            if (HasSchema)
            {
                return ImportInternal(table, (property, colIndex) =>
                {
                    var name = PropertyNamePolicy == null ? property.Name : PropertyNamePolicy(property);
                    return name == table.Data[0][colIndex]?.ToString();
                });
            }
            else
            {
                return ImportInternal(table, (property, colIndex) =>
                {
                    var index = PropertyIndexPolicy == null ? Array.IndexOf(TProperties, property) : PropertyIndexPolicy(property);
                    return index == colIndex;
                });
            }
        }

        private ImportResult<IEnumerable<T>> ImportInternal(ExcelyTable table, Func<PropertyInfo, int, bool> propertyMatcher)
        {
            var result = new List<T>(table.MaxRowCount);
            var rowErrors = new Dictionary<CellLocation, Exception>();

            for (var row = HasSchema ? 1 : 0; row < table.MaxRowCount; row++)
            {
                var obj = new T();
                var rowParseSuccess = true;

                for (var col = 0; col < table.MaxColCount; col++)
                {
                    var property = TProperties.FirstOrDefault(p => propertyMatcher(p, col));

                    if (property != null)
                    {
                        var value = CustomValuePolicy != null
                            ? CustomValuePolicy(property, table.Data[row][col])
                            : table.Data[row][col];

                        try
                        {
                            property.SetValue(obj, value);
                        }
                        catch (Exception ex)
                        {
                            if (StopWhenError) throw;
                            rowErrors[new CellLocation(row, col)] = ex;
                            rowParseSuccess = false;
                        }
                    }
                }

                if (rowParseSuccess) result.Add(obj);
            }

            return new ImportResult<IEnumerable<T>>(result, rowErrors);
        }
    }
}
