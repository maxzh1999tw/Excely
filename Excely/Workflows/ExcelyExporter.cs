using Excely.Shaders;
using Excely.TableFactories;
using System.Reflection;

namespace Excely.Workflows
{
    /// <summary>
    /// 提供從資料輸入到匯出的完整工作流程。
    /// </summary>
    /// <typeparam name="TInput">資料的輸入型別</typeparam>
    public class ExcelyExporter<TInput> : BaseWorkflow<TInput>
    {
        /// <summary>
        /// 取得匯出結果後依序執行的 IShaders
        /// </summary>
        public IEnumerable<IShader> Shaders { get; set; } = Enumerable.Empty<IShader>();

        /// <summary>
        /// 以指定的 ITableFactory 物件初始化新的實例。
        /// </summary>
        public ExcelyExporter(ITableFactory<TInput> tableFactory) : base(tableFactory) { }
    }

    /// <summary>
    /// 提供用於快速建立 ExcelyExporter 的靜態方法。
    /// </summary>
    public static class ExcelyExporter
    {
        /// <summary>
        /// 建立一個以 Class list 為輸入資料型別的 ExcelyExporter。
        /// </summary>
        /// <typeparam name="TClass">欲轉換的 Class</typeparam>
        public static ExcelyExporter<IEnumerable<TClass>> FromClassList<TClass>(
            Func<PropertyInfo, bool>? propertyShowPolicy = null,
            Func<PropertyInfo, string?>? propertyNamePolicy = null,
            Func<PropertyInfo, int>? propertyOrderPolicy = null,
            Func<TClass, PropertyInfo, object?>? customValuePolicy = null
            ) where TClass : class
        {
            ITableFactory<IEnumerable<TClass>> tableFactory = new ClassListTableFactory<TClass>()
            {
                PropertyShowPolicy = propertyShowPolicy,
                PropertyNamePolicy = propertyNamePolicy,
                PropertyOrderPolicy = propertyOrderPolicy,
                CustomValuePolicy = customValuePolicy,
            };

            return new ExcelyExporter<IEnumerable<TClass>>(tableFactory);
        }
    }
}
