using Excely.Plugins;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excely.EPPlus.LGPL.Plugins
{
    public static class ErrorRecordPluginExtention
    {
        public static ExcelWorksheet ExportToWorksheet(this ErrorRecordPlugin errorRecordPlugin, bool errorRowOnly)
        {
            if(errorRecordPlugin?.Table == null)
            {
                throw new ArgumentNullException();
            }

            for(int rowIndex = 0; rowIndex < errorRecordPlugin.Table.MaxRowCount; rowIndex++)
            {
                if(!errorRecordPlugin.Errors.Keys.Any(cell => cell.Row == row))
            }
        }
    }
}
