using Excely.EPPlus.LGPL.TableFactories;
using OfficeOpenXml;

namespace Excely.Workflows
{
    /// <summary>
    /// 提供快速建立從 Excel 到指定資料結構之 Importer 的方法。
    /// </summary>
    public class XlsxImporterBuilder : ExcelyImporterBuilder<ExcelWorksheet>
    {
        public XlsxImporterBuilder() : base(new XlsxTableFactory()) { }
    }
}
