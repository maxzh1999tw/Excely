﻿using Excely.Shaders;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excely.EPPlus.LGPL.Shaders.Xlsx
{
    /// <summary>
    /// 將出錯的表格醒目標示，並標註錯誤原因。
    /// </summary>
    public class ErrorMarkShader : XlsxShaderBase
    {
        /// <summary>
        /// 匯入錯誤資料。
        /// </summary>
        public Dictionary<CellLocation, string> CellErrors { get; set; }

        /// <summary>
        /// 醒目標示文字顏色。
        /// </summary>
        public Color TextColor { get; set; } = Color.Red;

        /// <summary>
        /// Comment 作者(不得為空)。
        /// </summary>
        public string Author { get; set; }

        /// <param name="errorCells">匯入錯誤資料</param>
        /// <param name="author">Comment 作者(不得為空)</param>
        public ErrorMarkShader(Dictionary<CellLocation, string> errorCells, string author)
        {
            if (string.IsNullOrEmpty(author))
            {
                throw new ArgumentException($"{nameof(author)} 參數不得為空。");
            }

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
