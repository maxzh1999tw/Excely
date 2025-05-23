using ClosedXML.Excel;
using Excely.Shaders;

namespace Excely.ClosedXML.Shaders
{
    /// <summary>
    /// Xlsx 專用的 Shader 開發基礎類別。
    /// </summary>
    public abstract class XlsxShaderBase : IShader
    {
        public T Execute<T>(T target)
        {
            if (target is IXLWorksheet worksheet)
            {
                ExecuteOnWorksheet(worksheet);
                return target;
            }

            throw new NotImplementedException();
        }

        protected abstract void ExecuteOnWorksheet(IXLWorksheet worksheet);
    }
}
