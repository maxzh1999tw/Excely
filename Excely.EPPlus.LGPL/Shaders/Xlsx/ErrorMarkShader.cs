using Excely.Shaders;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excely.EPPlus.LGPL.Shaders.Xlsx
{
    public class ErrorMarkShader : XlsxShaderBase
    {
        public Dictionary<CellLocation, string> CellErrors { get; set; }
        public Color TextColor { get; set; } = Color.Red;
        public string Author { get; set; }

        public ErrorMarkShader(Dictionary<CellLocation, string> errorCells, string author)
        {
            CellErrors = errorCells;
            Author = author;
        }

        protected override void ExcuteOnWorksheet(ExcelWorksheet worksheet)
        {
            foreach (var cellError in CellErrors)
            {
                var cell = worksheet.Cells[cellError.Key.Row+1, cellError.Key.Column+1];
                cell.Style.Font.Color.SetColor(TextColor);
                cell.AddComment(cellError.Value, Author);
            }
        }
    }
}
