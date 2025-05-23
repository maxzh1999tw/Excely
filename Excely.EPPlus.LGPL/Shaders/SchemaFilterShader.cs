using OfficeOpenXml;

namespace Excely.EPPlus.LGPL.Shaders
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

        protected override void ExecuteOnWorksheet(ExcelWorksheet target)
        {
            var schemaLength = SchemaLength;
            if (schemaLength == 0)
            {
                schemaLength = target.Dimension.End.Column;
            }
            target.Cells[StartCell.Row + 1, StartCell.Column + 1, StartCell.Row + 1, StartCell.Column + schemaLength].AutoFilter = true;
        }
    }
}
