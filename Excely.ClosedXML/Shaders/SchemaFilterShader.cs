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

        /// <param name="schemaLength">表頭長度</param>
        public SchemaFilterShader(int schemaLength)
        {
            SchemaLength = schemaLength;
        }

        /// <param name="schemaLength">表頭長度</param>
        /// <param name="startCell">表格起始儲存格座標</param>
        public SchemaFilterShader(int schemaLength, CellLocation startCell)
        {
            SchemaLength = schemaLength;
            StartCell = startCell;
        }

        protected override void ExcuteOnWorksheet(IXLWorksheet target)
        {
            target.Range(StartCell.Row + 1, StartCell.Column + 1, StartCell.Row + 1, StartCell.Column + SchemaLength).SetAutoFilter();
        }
    }
}
