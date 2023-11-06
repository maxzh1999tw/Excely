using Excely.TableFactories;
using Excely.TableImporter;
using System.Reflection;

namespace Excely.Workflows
{
    /// <summary>
    /// 建立 Importer 時的中繼類別，
    /// 已經指定了資料來源轉換為 ExcelyTable 的 TableFactory，
    /// 提供指定匯入結構並建立 Importer 的函式。 
    /// </summary>
    /// <typeparam name="TInput">資料輸入型別</typeparam>
    public abstract class ExcelyImporterBuilder<TInput>
    {
        /// <summary>
        /// 用於將輸入資料轉換為 ExcelyTable 的物件。
        /// </summary>
        public ITableFactory<TInput> TableFactory { get; set; }

        public ExcelyImporterBuilder(ITableFactory<TInput> tableFactory)
        {
            TableFactory = tableFactory;
        }

        /// <summary>
        /// 建立目標匯入型別為 Class list 的 Importer
        /// </summary>
        /// <typeparam name="TClass">目標 Class</typeparam>
        /// <param name="hasSchema">
        /// 匯入的 Table 是否含有表頭。
        /// 當此欄位為 true 時，會使用 PropertyNamePolicy；
        /// 否則會使用 PropertyIndexPolicy。
        /// </param>
        /// <param name="stopWhenError">
        /// 當轉換發生錯誤時是否立刻停止。
        /// 若此欄為 false，則發生錯誤時會跳過該 Row，繼續執行匯入。
        /// </param>
        /// <param name="propertyNamePolicy">
        /// 決定 Property 作為欄位時的名稱。
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位名稱」，
        /// 若此屬性為 null，則欄位名稱為 PropertyInfo.Name。
        /// </param>
        /// <param name="propertyIndexPolicy">
        /// 決定 Property 出現在表頭時的位置。
        /// 輸入參數為 PropertyInfo，輸出結果為「欄位索引」，
        /// 若該 Property 沒有出現在表頭中，請回傳 null。
        /// 若此屬性為 null，則依照 property 預設順序排序。
        /// </param>
        /// <param name="customValuePolicy">
        /// 決定將值寫入至 Property 時應寫入的值。
        /// 輸入參數為 (PropertyInfo, 原始值)，輸出結果為「應寫入的值」。
        /// </param>
        /// <param name="errorHandlingPolicy">
        /// 將值輸入進物件發生錯誤時，決定錯誤處理方式。
        /// 輸入參數為 (儲存格座標, 目標物件, PropertyInfo, 嘗試輸入的值, 發生的錯誤, 是否忽略此錯誤)
        /// 若此屬性為 null，則一律拋出異常。
        /// </param>
        /// <returns></returns>
        public ExcelyImporter<TInput, IEnumerable<TClass>> BuildForClassList<TClass>(
            bool hasSchema = true, bool stopWhenError = true,
            Func<PropertyInfo, string?>? propertyNamePolicy = null,
            Func<PropertyInfo, int?>? propertyIndexPolicy = null,
            Func<PropertyInfo, object?, object?>? customValuePolicy = null,
            Func<CellLocation, TClass, PropertyInfo, object?, Exception, bool>? errorHandlingPolicy = null)
            where TClass : class, new()
        {
            return new ExcelyImporter<TInput, IEnumerable<TClass>>(TableFactory, new ClassListTableImporter<TClass>()
            {
                HasSchema = hasSchema,
                StopWhenError = stopWhenError,
                PropertyNamePolicy = propertyNamePolicy,
                PropertyIndexPolicy = propertyIndexPolicy,
                CustomValuePolicy = customValuePolicy,
                ErrorHandlingPolicy = errorHandlingPolicy
            });
        }
    }
}
