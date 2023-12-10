using ClosedXML.Excel;
using Excely.ClosedXML.Shaders;
using Excely.ClosedXML.Workflows;
using Excely.Example.Console.Models;
using Excely.Example.Console.Utilities;
using Excely.Shaders;
using Excely.TableConverters;
using Excely.TableFactories;
using Excely.Workflows;

namespace Excely.Example.Console.Exemples
{
    internal class XlsxImportErrorHandlingExemple
    {
        private static IEnumerable<SimpleClass> GetList()
        {
            return new List<SimpleClass>()
            {
                new SimpleClass()
                {
                    Id = 1,
                    StringField = "比較長的文字欄位111111111111111111",
                    DateTimeField = DateTime.Now,
                    BoolField = true,
                },
                new SimpleClass()
                {
                    Id = 2,
                    StringField = "比較長的文字欄位222222222222222222",
                    DateTimeField = DateTime.Now,
                    BoolField = false,
                },
            };
        }

        private static XLWorkbook GetExampleWorkbook()
        {
            var list = GetList();
            var exportOption = new ClassListTableFactoryOptions<SimpleClass>
            {
                PropertyNamePolicy = p => p.GetDisplayName(),
                CustomValuePolicy = (p, obj) => p.Name switch
                {
                    // 故意設定錯誤的日期格式
                    nameof(SimpleClass.DateTimeField) => obj.DateTimeField?.ToString("yyyy/yyyy/dd"),
                    nameof(SimpleClass.BoolField) => obj.BoolField == null ? null : (obj.BoolField.Value ? "是" : "否"),
                    _ => p.GetValue(obj)
                },
            };

            var shaders = new IShader[]
            {
                new SchemaFilterShader(),
                new CellFittingShader(),
            };

            var exporter = ExcelyExporter.FromClassList(exportOption, shaders);
            return exporter.ToExcel(list);
        }

        public static void Demo()
        {
            using var excel = GetExampleWorkbook();
            var worksheet = excel.Worksheets.First();
            var importer = new XlsxImporter();
            var errorDict = new Dictionary<CellLocation, string>();

            var importOption = new ClassListTableConverterOptions<SimpleClass>
            {
                PropertyNamePolicy = p => p.GetDisplayName(),
                PropertyValueSettingPolicy = (p, value) => p.Name switch
                {
                    nameof(SimpleClass.BoolField) => value != null ? value?.ToString() == "是" : null,
                    _ => value
                },
                ErrorHandlingPolicy = (cellLocation, obj, p, value, ex) =>
                {
                    switch (p.Name)
                    {
                        case nameof(SimpleClass.DateTimeField):
                         errorDict.Add(cellLocation, "錯誤的日期格式");
                            break;
                        default:
                            errorDict.Add(cellLocation, ex.Message);
                            break;
                    };
                    return true;
                }
            };

            var importedList = importer.ToClassList(worksheet, importOption);
            new ErrorMarkShader(errorDict).Excute(worksheet);
            new CellFittingShader().Excute(worksheet);
            new SchemaFilterShader().Excute(worksheet);
            new TableThemeShader().Excute(worksheet);
            excel.SaveAs("XlsxImportErrorHandlingExemple.xlsx");
        }
    }
}
