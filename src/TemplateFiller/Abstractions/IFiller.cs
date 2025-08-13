namespace TemplateFiller.Abstractions
{
    /// <summary>
    /// 表示一个数据填充器
    /// </summary>
    public interface IFiller
    {
        /// <summary>
        /// 从<paramref name="source"/>中读取与占位符路径匹配的数据，并填充到目标
        /// </summary>
        /// <param name="source"></param>
        public void Fill(ISource source);
        /// <summary>
        /// 检查目标是否有匹配的占位符
        /// </summary>
        /// <returns>如果有匹配的占位符，则返回true；否则返回false</returns>
        public bool Check();
    }
}
