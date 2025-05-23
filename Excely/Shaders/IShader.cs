namespace Excely.Shaders
{
    /// <summary>
    /// 定義了一個可以調整匯出結果的執行單元。
    /// </summary>
    public interface IShader
    {
        /// /// <summary>
        /// 對特定匯出結果進行調整。
        /// </summary>
        /// <param name="target">執行目標</param>
        /// <typeparam name="TOutput">匯出結果的載體</typeparam>
        /// <returns>執行結果</returns>
        TOutput Execute<TOutput>(TOutput target);
    }
}
