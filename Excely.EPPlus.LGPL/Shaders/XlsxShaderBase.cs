using Excely.Shaders;
using OfficeOpenXml;

namespace Excely.EPPlus.LGPL.Shaders
{
    /// <summary>
    /// Xlsx 專用的 Shader 開發基礎類別。
    /// </summary>
    public abstract class XlsxShaderBase : IShader
    {
        public T Excute<T>(T target)
        {
            if (target is ExcelWorksheet worksheet)
            {
                ExcuteOnWorksheet(worksheet);
                return target;
            }

            throw new NotImplementedException();
        }

        protected abstract void ExcuteOnWorksheet(ExcelWorksheet worksheet);
    }
}
