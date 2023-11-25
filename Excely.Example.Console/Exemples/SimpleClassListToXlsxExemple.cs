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
    internal class SimpleClassListToXlsxExemple
    {
        public static void Demo()
        {
            #region === 匯出 ===
            var list = GetList();
            var exportOption = new ClassListTableFactoryOptions<SimpleClass>
            {
                PropertyNamePolicy = p => p.GetDisplayName(),
                CustomValuePolicy = (p, obj) => p.Name switch
                {
                    nameof(SimpleClass.DateTimeField) => obj.DateTimeField?.ToString("yyyy/MM/dd"),
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
            using var excel = exporter.ToExcel(list);
            excel.SaveAs("SimpleClassListToXlsxExemple.xlsx");
            #endregion

            #region === 匯入 ===
            var worksheet = excel.Worksheets.First();
            var importer = new XlsxImporter();

            var importOption = new ClassListTableConverterOptions<SimpleClass>
            {
                PropertyNamePolicy = p => p.GetDisplayName(),
                PropertyValueSettingPolicy = (p, value) => p.Name switch
                {
                    nameof(SimpleClass.DateTimeField) => value != null ? DateTime.Parse(value.ToString()) : null,
                    nameof(SimpleClass.BoolField) => value != null ? value == "是" : null,
                    _ => value
                },
            };

            var importedList = importer.ToClassList(worksheet, importOption);
            #endregion
        }

        private static IEnumerable<SimpleClass> GetList()
        {
            return new List<SimpleClass>()
            {
                new SimpleClass()
                {
                    Id = 1,
                    StringField = "Text",
                    DateTimeField = DateTime.Now,
                    BoolField = true,
                },
                new SimpleClass()
                {
                    Id = 2,
                    StringField = null,
                    DateTimeField = null,
                    BoolField = null,
                },
            };
        }
    }
}
