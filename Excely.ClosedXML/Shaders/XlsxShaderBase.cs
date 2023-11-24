using ClosedXML.Excel;
using Excely.Shaders;

namespace Excely.ClosedXML.Shaders
{
    /// <summary>
    /// Xlsx 專用的 Shader 開發基礎類別。
    /// </summary>
    public abstract class XlsxShaderBase : IShader
    {
        public T Excute<T>(T target)
        {
            if (target is IXLWorksheet worksheet)
            {
                ExcuteOnWorksheet(worksheet);
                return target;
            }

            throw new NotImplementedException();
        }

        protected abstract void ExcuteOnWorksheet(IXLWorksheet worksheet);
    }
}
