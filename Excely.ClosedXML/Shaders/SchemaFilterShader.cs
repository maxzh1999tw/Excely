using ClosedXML.Excel;

namespace Excely.ClosedXML.Shaders
{
    /// <summary>
    /// 為表頭啟用篩選。
    /// </summary>
    public class SchemaFilterShader : XlsxShaderBase
    {
        /// <summary>
        /// 表頭寬度。
        /// </summary>
        public int SchemaLength { get; set; }

        /// <summary>
        /// 表格起始儲存格座標。
        /// </summary>
        public CellLocation StartCell { get; set; } = new(0, 0);

        public SchemaFilterShader() { }

        /// <param name="schemaLength">表頭長度(0 為自適應)</param>
        /// <param name="startCell">表格起始儲存格座標</param>
        public SchemaFilterShader(int schemaLength, CellLocation startCell)
        {
            SchemaLength = schemaLength;
            StartCell = startCell;
        }

        protected override void ExcuteOnWorksheet(IXLWorksheet target)
        {
            var schemaLength = SchemaLength;
            if (SchemaLength == 0)
            {
                schemaLength = target.LastColumnUsed().ColumnNumber();
            }
            target.Range(StartCell.Row + 1, StartCell.Column + 1, StartCell.Row + 1, StartCell.Column + schemaLength).SetAutoFilter();
        }
    }
}
