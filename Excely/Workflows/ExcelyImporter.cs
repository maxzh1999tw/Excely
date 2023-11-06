using Excely.TableFactories;
using Excely.TableConverters;

namespace Excely.Workflows
{
    /// <summary>
    /// 提供從資料輸入到轉換為資料結構的完整工作流程。
    /// </summary>
    /// <typeparam name="TInput">資料的輸入型別</typeparam>
    public abstract class ExcelyImporter<TInput>
    {
        public abstract ExcelyTable GetTable(TInput input);

        /// <summary>
        /// 將指定的資料來源匯入為指定資料結構
        /// </summary>
        /// <param name="dataSource">資料來源</param>
        /// <returns>匯入結果</returns>
        public IEnumerable<TClass> ToClassList<TClass>(
            TInput dataSource, 
            ClassListTableConverter<TClass>? converter = null) 
            where TClass : class, new()
        {
            var table = GetTable(dataSource);
            converter ??= new ClassListTableConverter<TClass>();
            return converter.Convert(table);
        }
    }
}

