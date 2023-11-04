﻿using Excely.Shaders;
using Excely.TableFactories;

namespace Excely.Exporters
{
    /// <summary>
    /// 提供從資料輸入到匯出的完整工作流程。
    /// </summary>
    /// <typeparam name="TInput">資料的輸入型別</typeparam>
    public class ExcelyExporter<TInput>
    {
        /// <summary>
        /// 將輸入資料轉換為 Table 的 ITableFactory 物件。
        /// </summary>
        protected ITableFactory<TInput> TableFactory { get; set; }

        /// <summary>
        /// 取得匯出結果後依序執行的 IShaders
        /// </summary>
        public IEnumerable<IShader> Shaders { get; set; } = Enumerable.Empty<IShader>();

        /// <summary>
        /// 以指定的 ITableFactory 物件初始化新的實例。
        /// </summary>
        public ExcelyExporter(ITableFactory<TInput> tableFactory)
        {
            TableFactory = tableFactory;
        }

        /// <summary>
        /// 將輸入資料轉換為 Table。
        /// </summary>
        /// <param name="sourceData">輸入資料</param>
        /// <returns>ExcelyTable 格式的資料</returns>
        public ExcelyTable GetTable(TInput sourceData) => TableFactory.GetTable(sourceData);
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
        public static ExcelyExporter<IEnumerable<TClass>> FromClassList<TClass>() where TClass : class
        {
            ITableFactory<IEnumerable<TClass>> tableFactory = new ClassListTableFactory<TClass>();

            return new ExcelyExporter<IEnumerable<TClass>>(tableFactory);
        }
    }
}
